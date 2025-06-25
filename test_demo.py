"""
Excel分析ツールの使用例とテストスクリプト
"""

import os
from excel_utils import analyze_excel_structure, excel_to_text_representation
from llm_api import analyze_excel_with_llm

def test_excel_analysis():
    """サンプルファイルを使った分析テスト"""
    
    # サンプルファイルのパス
    sample_files = [
        "sample_single_header.xlsx",
        "sample_multi_header.xlsx", 
        "sample_form.xlsx"
    ]
    
    print("=== Excel分析ツール テスト ===\n")
    
    for filename in sample_files:
        if not os.path.exists(filename):
            print(f"ファイルが見つかりません: {filename}")
            continue
            
        print(f"📊 分析中: {filename}")
        print("-" * 50)
        
        # ファイルを読み込み
        with open(filename, 'rb') as f:
            file_content = f.read()
        
        # 基本構造分析
        print("【基本構造分析】")
        basic_analysis = analyze_excel_structure(file_content)
        
        if "error" in basic_analysis:
            print(f"エラー: {basic_analysis['error']}")
            continue
            
        for sheet_name, stats in basic_analysis.items():
            print(f"シート: {sheet_name}")
            print(f"  行数: {stats['max_row']}")
            print(f"  列数: {stats['max_column']}")
            print(f"  データ密度: {stats['data_density']:.1%}")
            print(f"  結合セル: {stats['num_merged_cells']}個")
            print(f"  Excelテーブル: {'あり' if stats['has_excel_tables'] else 'なし'}")
        
        # テキスト表現の生成
        print("\n【テキスト表現】")
        text_representations = excel_to_text_representation(file_content, max_rows=10)
        
        if "error" in text_representations:
            print(f"エラー: {text_representations['error']}")
            continue
            
        for sheet_name, text_repr in text_representations.items():
            print(f"シート: {sheet_name}")
            print(text_repr[:300] + "..." if len(text_repr) > 300 else text_repr)
        
        print("\n" + "="*70 + "\n")

def demo_llm_analysis():
    """LLM分析のデモ（APIキーが必要）"""
    
    print("=== LLM分析デモ ===")
    print("注意: このデモを実行するには、OpenAI APIキーが必要です。")
    print("環境変数 OPENAI_API_KEY を設定するか、直接入力してください。\n")
    
    # APIキーの取得
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        api_key = input("OpenAI APIキーを入力してください: ").strip()
    
    if not api_key:
        print("APIキーが入力されていません。デモを終了します。")
        return
    
    # サンプルファイルで分析
    filename = "sample_single_header.xlsx"
    
    if not os.path.exists(filename):
        print(f"サンプルファイルが見つかりません: {filename}")
        return
    
    print(f"📊 LLM分析中: {filename}")
    
    with open(filename, 'rb') as f:
        file_content = f.read()
    
    # テキスト表現に変換
    text_representations = excel_to_text_representation(file_content)
    
    if "error" in text_representations:
        print(f"エラー: {text_representations['error']}")
        return
    
    # LLM分析を実行
    print("LLMに送信中...")
    analysis_result = analyze_excel_with_llm(text_representations, api_key)
    
    if analysis_result:
        print("✅ 分析完了!")
        print("\n【分析結果】")
        
        for sheet_result in analysis_result.sheets:
            print(f"シート: {sheet_result.sheet_name}")
            print(f"分類: {sheet_result.sheet_type}")
            
            if sheet_result.header_info:
                print(f"ヘッダー: 行{sheet_result.header_info.start_row}-{sheet_result.header_info.end_row} ({sheet_result.header_info.header_type})")
            else:
                print("ヘッダー: なし")
            
            print(f"理由: {sheet_result.reasoning}")
            
            # pandas読み込み例
            if sheet_result.header_info and sheet_result.sheet_type in ["table", "mixed"]:
                print(f"\n【pandas読み込み例】")
                print(f"df = pd.read_excel('{filename}', sheet_name='{sheet_result.sheet_name}', header={list(range(sheet_result.header_info.start_row - 1, sheet_result.header_info.end_row))})")
    else:
        print("❌ 分析に失敗しました。APIキーを確認してください。")

if __name__ == "__main__":
    print("Excel分析ツール テストスクリプト")
    print("1. 基本分析テスト")
    print("2. LLM分析デモ")
    print("3. 両方実行")
    
    choice = input("\n選択してください (1/2/3): ").strip()
    
    if choice in ["1", "3"]:
        test_excel_analysis()
    
    if choice in ["2", "3"]:
        demo_llm_analysis()
    
    print("\nテスト完了!")

