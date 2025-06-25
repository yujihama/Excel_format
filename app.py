"""
Excel分析Streamlitアプリケーション
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import json

from models import ExcelAnalysisOutput
from excel_utils import analyze_excel_structure, excel_to_text_representation, format_analysis_results
from llm_api import analyze_excel_with_llm, test_llm_functionality


def main():
    """メインアプリケーション"""
    st.set_page_config(
        page_title="Excel構造分析ツール",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("Excel構造分析ツール")
    st.markdown("Excelファイルをアップロードして、テーブル形式の判別とヘッダー検出を行います。")
    
    # サイドバーでAPIキー入力
    with st.sidebar:
        st.header("設定")
        api_key = st.text_input(
            "OpenAI APIキー",
            type="password",
            help="OpenAI APIキーを入力してください"
        )
        
        if api_key:
            st.success("APIキーが設定されました")
            
            # APIキーのテスト
            if st.button("APIキーをテスト"):
                with st.spinner("APIキーをテスト中..."):
                    success, message = test_llm_functionality(api_key)
                    if success:
                        st.success(message)
                    else:
                        st.error(message)
        else:
            st.warning("OpenAI APIキーを入力してください")
    
    # メインコンテンツ
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("ファイルアップロード")
        
        uploaded_file = st.file_uploader(
            "Excelファイルを選択してください",
            type=['xlsx', 'xls'],
            help="対応形式: .xlsx, .xls"
        )
        
        if uploaded_file is not None:
            st.success(f"ファイル '{uploaded_file.name}' がアップロードされました")
            
            # ファイル情報を表示
            file_details = {
                "ファイル名": uploaded_file.name,
                "ファイルサイズ": f"{uploaded_file.size:,} bytes"
            }
            st.json(file_details)
            
            # 基本構造分析
            with st.expander("基本構造分析", expanded=True):
                file_content = uploaded_file.read()
                uploaded_file.seek(0)  # ファイルポインタをリセット
                
                basic_analysis = analyze_excel_structure(file_content)
                
                if "error" in basic_analysis:
                    st.error(basic_analysis["error"])
                else:
                    st.subheader("シート別統計")
                    for sheet_name, stats in basic_analysis.items():
                        with st.container():
                            st.write(f"**{sheet_name}**")
                            col_a, col_b, col_c = st.columns(3)
                            with col_a:
                                st.metric("行数", stats['max_row'])
                            with col_b:
                                st.metric("列数", stats['max_column'])
                            with col_c:
                                st.metric("データ密度", f"{stats['data_density']:.1%}")
                            
                            if stats['num_merged_cells'] > 0:
                                st.info(f"結合セル: {stats['num_merged_cells']}個")
                            if stats['has_excel_tables']:
                                st.info(f"Excelテーブル: {stats['num_excel_tables']}個")
    
    with col2:
        st.header("LLM分析結果")
        
        if uploaded_file is not None and api_key:
            # 画像キャプチャ機能の設定
            st.subheader("分析オプション")
            use_image_capture = st.checkbox("LibreOfficeで画像キャプチャを取得して分析精度を向上", value=True)
            
            if st.button("LLM分析を実行", type="primary"):
                with st.spinner("分析中... しばらくお待ちください"):
                    try:
                        # ファイル内容を再読み込み
                        file_content = uploaded_file.read()
                        uploaded_file.seek(0)
                        
                        # 画像キャプチャを取得（オプション）
                        image_paths = None
                        if use_image_capture:
                            with st.spinner("LibreOfficeで画像キャプチャを取得中..."):
                                from libreoffice_capture import capture_excel_with_libreoffice
                                image_paths = capture_excel_with_libreoffice(file_content, "/tmp/excel_captures")
                                
                                if image_paths:
                                    st.success(f"画像キャプチャを取得しました ({len(image_paths)}枚)")
                                    
                                    # 取得した画像を表示
                                    with st.expander("取得した画像キャプチャ"):
                                        for i, img_path in enumerate(image_paths):
                                            st.image(img_path, caption=f"シート {i+1}", use_column_width=True)
                                else:
                                    st.warning("画像キャプチャの取得に失敗しました。テキスト分析のみ実行します。")
                        
                        # テキスト表現に変換
                        text_representations = excel_to_text_representation(file_content)
                        
                        if "error" in text_representations:
                            st.error(text_representations["error"])
                        else:
                            # LLM分析を実行（画像付きまたはテキストのみ）
                            analysis_result = analyze_excel_with_llm(text_representations, api_key, image_paths=image_paths)
                            
                            if analysis_result:
                                st.success("分析が完了しました")
                                
                                # 分析方法を表示
                                if image_paths:
                                    st.info("テキスト情報と画像キャプチャの両方を使用して分析しました")
                                else:
                                    st.info("テキスト情報のみを使用して分析しました")
                                
                                # 結果を表示
                                formatted_result = format_analysis_results(analysis_result)
                                st.markdown(formatted_result)
                                
                                # 詳細結果をJSON形式で表示
                                with st.expander("詳細結果 (JSON)"):
                                    st.json(analysis_result.model_dump())
                                
                                # pandasでの読み込み例を提示
                                st.subheader("pandas読み込み例")
                                for sheet_result in analysis_result.sheets:
                                    if sheet_result.header_info and sheet_result.sheet_type in ["table", "mixed"]:
                                        st.code(f"""
# シート '{sheet_result.sheet_name}' の読み込み例
import pandas as pd

# ヘッダー行: {sheet_result.header_info.start_row}-{sheet_result.header_info.end_row}
df = pd.read_excel(
    'your_file.xlsx',
    sheet_name='{sheet_result.sheet_name}',
    header={list(range(sheet_result.header_info.start_row - 1, sheet_result.header_info.end_row))}
)
""", language="python")
                            else:
                                st.error("分析に失敗しました。APIキーを確認してください。")
                    
                    except Exception as e:
                        st.error(f"エラーが発生しました: {str(e)}")
        
        elif uploaded_file is None:
            st.info("Excelファイルをアップロードしてください")
        elif not api_key:
            st.info("OpenAI APIキーを入力してください")
    
    # テキスト表現の表示（デバッグ用）
    if uploaded_file is not None:
        with st.expander("テキスト表現 (デバッグ用)"):
            file_content = uploaded_file.read()
            uploaded_file.seek(0)
            
            text_representations = excel_to_text_representation(file_content, max_rows=10)
            
            if "error" not in text_representations:
                for sheet_name, text_repr in text_representations.items():
                    st.subheader(f"シート: {sheet_name}")
                    st.text(text_repr)
    
    # フッター
    st.markdown("---")
    st.markdown(
        "このツールは、ExcelファイルをLLMで分析し、テーブル構造とヘッダー情報を自動検出します。"
    )


if __name__ == "__main__":
    main()

