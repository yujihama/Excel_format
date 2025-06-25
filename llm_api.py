"""
LLM API呼び出し機能
"""

import json
import openai
import os
import base64
from typing import Optional, Dict, Any, List
from pydantic import ValidationError

from models import ExcelAnalysisOutput


def create_analysis_prompt(excel_text_representation: str) -> str:
    """
    Excel分析用のプロンプトを作成する
    
    Args:
        excel_text_representation: Excelのテキスト表現
        
    Returns:
        str: LLM用のプロンプト
    """
    prompt = f"""あなたはExcelシートの構造分析に特化したAIアシスタントです。
あなたのタスクは、Excelシートの内容と構造に基づいて、シートを「table（テーブル）」「form（フォーム）」「mixed（混合）」「unknown（不明）」のいずれかに分類することです。
シートが「table」または「mixed」に分類された場合、ヘッダー行も特定してください。

以下のJSONスキーマに準拠したJSONオブジェクトとしてのみ出力してください。

```json
{{
  "type": "object",
  "properties": {{
    "sheets": {{
      "type": "array",
      "items": {{
        "type": "object",
        "properties": {{
          "sheet_name": {{ "type": "string", "description": "分析対象のExcelシート名。" }},
          "sheet_type": {{ "type": "string", "enum": ["table", "form", "mixed", "unknown"], "description": "シートの主要な目的の分類。" }},
          "header_info": {{
            "type": "object",
            "properties": {{
              "start_row": {{ "type": "integer", "description": "ヘッダーが始まる1ベースの行インデックス。" }},
              "end_row": {{ "type": "integer", "description": "ヘッダーが終わる1ベースの行インデックス。" }},
              "header_type": {{ "type": "string", "enum": ["single", "multi-level"], "description": "ヘッダーのタイプ: 'single' は単一行、'multi-level' は複数行。" }}
            }},
            "description": "検出されたテーブルヘッダーの詳細。テーブルでない場合やヘッダーが見つからない場合はnull。"
          }},
          "reasoning": {{ "type": "string", "description": "分類とヘッダー検出に関する簡潔な説明。" }}
        }},
        "required": ["sheet_name", "sheet_type", "reasoning"]
      }},
      "description": "Excelファイル内の各シートの分析結果。"
    }}
  }},
  "required": ["sheets"]
}}
```

分類の基準:
- **table**: データが行列形式で整理されており、明確なヘッダーがある
- **form**: 入力フォームのような構造で、ラベルと値のペアが散在している
- **mixed**: テーブル要素とフォーム要素が混在している
- **unknown**: 上記のいずれにも明確に分類できない

ヘッダー検出の基準:
- 太字(**テキスト**)で表示されているセル
- 結合セル([MERGED])を使用した階層構造
- データ行の上部に位置する説明的なテキスト

以下に、Excelシートの内容をテキスト形式で表現したものです：

{excel_text_representation}

JSONオブジェクトのみを出力してください。JSON以外のテキストや説明は含めないでください。"""
    
    return prompt


def call_openai_api_with_images(prompt: str, api_key: str, model: str, image_paths: List[str]) -> Optional[str]:
    """
    OpenAI APIを画像付きで呼び出してExcel分析を行う
    
    Args:
        prompt: 分析用プロンプト
        api_key: OpenAI APIキー
        model: 使用するモデル名
        image_paths: 画像ファイルのパスのリスト
        
    Returns:
        Optional[str]: APIからの応答（JSON文字列）
    """
    try:
        client = openai.OpenAI(api_key=api_key)
        
        # メッセージコンテンツを構築
        content = [
            {
                "type": "text",
                "text": prompt
            }
        ]
        
        # 画像を追加（最大3枚まで）
        for i, image_path in enumerate(image_paths[:3]):
            if os.path.exists(image_path):
                # 画像をBase64エンコード
                with open(image_path, "rb") as image_file:
                    base64_image = base64.b64encode(image_file.read()).decode('utf-8')
                
                content.append({
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{base64_image}",
                        "detail": "high"
                    }
                })
        
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "あなたはExcelファイルの構造分析を専門とするAIアシスタントです。テキスト情報と画像の両方を参考にして、指定されたJSONスキーマに厳密に従って応答してください。画像からは視覚的なレイアウト、フォーマット、ヘッダーの強調表示などの情報を読み取ってください。"
                },
                {
                    "role": "user",
                    "content": content
                }
            ],
            temperature=0.1,
            max_tokens=2000
        )
        
        return response.choices[0].message.content
        
    except openai.AuthenticationError as e:
        print(f"認証エラー: APIキーが無効です - {e}")
        return None
    except openai.RateLimitError as e:
        print(f"レート制限エラー: {e}")
        return None
    except openai.APIError as e:
        print(f"OpenAI APIエラー: {e}")
        return None
    except Exception as e:
        print(f"予期せぬエラー: {e}")
        return None


def call_openai_api(prompt: str, api_key: str, model: str = "gpt-4.1-mini") -> Optional[str]:
    """
    OpenAI APIを呼び出してExcel分析を行う
    
    Args:
        prompt: 分析用プロンプト
        api_key: OpenAI APIキー
        model: 使用するモデル名
        
    Returns:
        Optional[str]: APIからの応答（JSON文字列）
    """
    try:
        client = openai.OpenAI(api_key=api_key)
        
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "あなたはExcelファイルの構造分析を専門とするAIアシスタントです。指定されたJSONスキーマに厳密に従って応答してください。"
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.1,  # 一貫性のある結果を得るため低めに設定
            max_tokens=2000
        )
        
        return response.choices[0].message.content
        
    except openai.AuthenticationError as e:
        print(f"認証エラー: APIキーが無効です - {e}")
        return None
    except openai.RateLimitError as e:
        print(f"レート制限エラー: {e}")
        return None
    except openai.APIError as e:
        print(f"OpenAI APIエラー: {e}")
        return None
    except Exception as e:
        print(f"予期せぬエラー: {e}")
        return None


def analyze_excel_with_llm(excel_text_representations: Dict[str, str], api_key: str, model: str = "gpt-4.1-mini", image_paths: Optional[List[str]] = None) -> Optional[ExcelAnalysisOutput]:
    """
    LLMを使用してExcelファイルを分析する
    
    Args:
        excel_text_representations: シート名をキーとするテキスト表現の辞書
        api_key: OpenAI APIキー
        model: 使用するモデル名（デフォルト: gpt-4.1-mini）
        image_paths: Excelファイルのキャプチャ画像パスのリスト（オプション）
        
    Returns:
        Optional[ExcelAnalysisOutput]: 分析結果
    """
    if not excel_text_representations:
        return None
    
    # 複数シートの場合は結合してプロンプトを作成
    combined_text = ""
    for sheet_name, text_repr in excel_text_representations.items():
        combined_text += f"\\n\\n=== シート: {sheet_name} ===\\n{text_repr}"
    
    prompt = create_analysis_prompt(combined_text)
    
    # 画像がある場合はマルチモーダル分析を実行
    if image_paths and len(image_paths) > 0:
        llm_response = call_openai_api_with_images(prompt, api_key, model, image_paths)
    else:
        # OpenAI APIを呼び出し（テキストのみ）
        llm_response = call_openai_api(prompt, api_key, model)
    
    if not llm_response:
        return None
    
    try:
        # LLM応答からマークダウンコードブロックを除去
        cleaned_response = llm_response.strip()
        if cleaned_response.startswith("```json"):
            cleaned_response = cleaned_response[7:]  # "```json"を除去
        if cleaned_response.startswith("```"):
            cleaned_response = cleaned_response[3:]  # "```"を除去
        if cleaned_response.endswith("```"):
            cleaned_response = cleaned_response[:-3]  # 末尾の"```"を除去
        cleaned_response = cleaned_response.strip()
        
        # JSONをパース
        json_data = json.loads(cleaned_response)
        
        # Pydanticモデルでバリデーション
        analysis_result = ExcelAnalysisOutput(**json_data)
        return analysis_result
        
    except json.JSONDecodeError as e:
        print(f"JSONデコードエラー: {e}")
        print(f"クリーニング後のLLM応答: {cleaned_response}")
        return None
    except ValidationError as e:
        print(f"バリデーションエラー: {e}")
        print(f"LLM応答: {llm_response}")
        return None
    except Exception as e:
        print(f"予期せぬエラー: {e}")
        return None


def test_llm_functionality(api_key: str) -> tuple[bool, str]:
    """
    LLM機能のテスト
    
    Args:
        api_key: OpenAI APIキー
        
    Returns:
        tuple[bool, str]: (テスト成功の場合True, エラーメッセージまたは成功メッセージ)
    """
    if not api_key or not api_key.strip():
        return False, "APIキーが入力されていません"
    
    test_excel_text = """
row1: | **ID** | **名前** | **値** |
row2: | 1 | りんご | 100 |
row3: | 2 | みかん | 150 |
"""
    
    test_representations = {"TestSheet": test_excel_text}
    
    try:
        # まず簡単なAPI接続テストを実行
        client = openai.OpenAI(api_key=api_key)
        
        # 簡単なテストリクエストを送信
        test_response = client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[{"role": "user", "content": "Hello"}],
            max_tokens=10
        )
        
        if not test_response.choices:
            return False, "APIからの応答が空です"
        
        # 実際の分析テストを実行
        result = analyze_excel_with_llm(test_representations, api_key)
        
        if result and len(result.sheets) > 0:
            return True, "APIキーのテストが成功しました"
        else:
            return False, "Excel分析テストに失敗しました"
            
    except openai.AuthenticationError as e:
        return False, f"認証エラー: APIキーが無効です ({str(e)})"
    except openai.RateLimitError as e:
        return False, f"レート制限エラー: {str(e)}"
    except openai.APIError as e:
        return False, f"OpenAI APIエラー: {str(e)}"
    except Exception as e:
        return False, f"予期せぬエラー: {str(e)}"

