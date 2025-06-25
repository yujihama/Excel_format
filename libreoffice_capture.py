"""
LibreOfficeを使用してExcelファイルのキャプチャを取得する機能
"""

import subprocess
import os
import tempfile
import time
from typing import Optional, List
import base64


def capture_excel_with_libreoffice(file_content: bytes, output_dir: str = "/tmp") -> Optional[List[str]]:
    """
    LibreOfficeを使用してExcelファイルのキャプチャを取得する
    
    Args:
        file_content: Excelファイルのバイナリコンテンツ
        output_dir: 出力ディレクトリ
        
    Returns:
        Optional[List[str]]: 生成された画像ファイルのパスのリスト
    """
    try:
        # 一時ファイルを作成
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
            temp_file.write(file_content)
            temp_excel_path = temp_file.name
        
        # 出力ディレクトリを作成
        os.makedirs(output_dir, exist_ok=True)
        
        # LibreOfficeでExcelファイルを開いてPDFに変換
        pdf_output_path = os.path.join(output_dir, "excel_capture.pdf")
        
        # LibreOfficeのヘッドレスモードでPDF変換
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            temp_excel_path
        ]
        
        # Xvfbを使用してヘッドレス環境でLibreOfficeを実行
        env = os.environ.copy()
        env['DISPLAY'] = ':99'
        
        # Xvfbを起動
        xvfb_process = subprocess.Popen(['Xvfb', ':99', '-screen', '0', '1024x768x24'], 
                                       stdout=subprocess.DEVNULL, 
                                       stderr=subprocess.DEVNULL)
        time.sleep(2)  # Xvfbの起動を待つ
        
        try:
            # LibreOfficeでPDF変換を実行
            result = subprocess.run(cmd, env=env, capture_output=True, text=True, timeout=30)
            
            if result.returncode != 0:
                print(f"LibreOffice変換エラー: {result.stderr}")
                return None
            
            # 生成されたPDFファイルのパスを確認
            base_name = os.path.splitext(os.path.basename(temp_excel_path))[0]
            pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
            
            if not os.path.exists(pdf_path):
                print(f"PDFファイルが生成されませんでした: {pdf_path}")
                return None
            
            # PDFを画像に変換
            image_paths = convert_pdf_to_images(pdf_path, output_dir)
            
            return image_paths
            
        finally:
            # Xvfbプロセスを終了
            xvfb_process.terminate()
            xvfb_process.wait()
            
    except Exception as e:
        print(f"キャプチャ取得エラー: {e}")
        return None
    finally:
        # 一時ファイルを削除
        if os.path.exists(temp_excel_path):
            os.unlink(temp_excel_path)


def convert_pdf_to_images(pdf_path: str, output_dir: str) -> List[str]:
    """
    PDFファイルを画像に変換する
    
    Args:
        pdf_path: PDFファイルのパス
        output_dir: 出力ディレクトリ
        
    Returns:
        List[str]: 生成された画像ファイルのパスのリスト
    """
    try:
        from pdf2image import convert_from_path
        
        # PDFを画像に変換
        images = convert_from_path(pdf_path, dpi=150, first_page=1, last_page=3)  # 最初の3ページまで
        
        image_paths = []
        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, f"excel_sheet_{i+1}.png")
            image.save(image_path, 'PNG')
            image_paths.append(image_path)
        
        return image_paths
        
    except ImportError:
        print("pdf2imageライブラリがインストールされていません")
        # poppler-utilsを使用してPDFを画像に変換
        return convert_pdf_with_poppler(pdf_path, output_dir)
    except Exception as e:
        print(f"PDF変換エラー: {e}")
        return []


def convert_pdf_with_poppler(pdf_path: str, output_dir: str) -> List[str]:
    """
    poppler-utilsを使用してPDFを画像に変換する
    
    Args:
        pdf_path: PDFファイルのパス
        output_dir: 出力ディレクトリ
        
    Returns:
        List[str]: 生成された画像ファイルのパスのリスト
    """
    try:
        # pdftoppmコマンドを使用してPDFを画像に変換
        output_prefix = os.path.join(output_dir, "excel_sheet")
        
        cmd = [
            "pdftoppm",
            "-png",
            "-f", "1",  # 最初のページから
            "-l", "3",  # 最大3ページまで
            "-r", "150",  # 解像度150 DPI
            pdf_path,
            output_prefix
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            print(f"pdftoppm変換エラー: {result.stderr}")
            return []
        
        # 生成された画像ファイルを検索
        image_paths = []
        for i in range(1, 4):  # 最大3ページ
            image_path = f"{output_prefix}-{i:02d}.png"
            if os.path.exists(image_path):
                image_paths.append(image_path)
        
        return image_paths
        
    except Exception as e:
        print(f"poppler変換エラー: {e}")
        return []


def encode_image_to_base64(image_path: str) -> Optional[str]:
    """
    画像ファイルをBase64エンコードする
    
    Args:
        image_path: 画像ファイルのパス
        
    Returns:
        Optional[str]: Base64エンコードされた画像データ
    """
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except Exception as e:
        print(f"画像エンコードエラー: {e}")
        return None


def test_libreoffice_capture():
    """
    LibreOfficeキャプチャ機能のテスト
    """
    print("LibreOfficeキャプチャ機能のテスト開始...")
    
    # サンプルファイルでテスト
    sample_file = "/home/ubuntu/excel_analyzer/sample_single_header.xlsx"
    
    if not os.path.exists(sample_file):
        print(f"サンプルファイルが見つかりません: {sample_file}")
        return False
    
    with open(sample_file, 'rb') as f:
        file_content = f.read()
    
    # キャプチャを取得
    image_paths = capture_excel_with_libreoffice(file_content, "/tmp/excel_captures")
    
    if image_paths:
        print(f"キャプチャ成功: {len(image_paths)}個の画像を生成")
        for path in image_paths:
            print(f"  - {path}")
        return True
    else:
        print("キャプチャ失敗")
        return False


if __name__ == "__main__":
    test_libreoffice_capture()

