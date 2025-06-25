"""
Excelファイル分析のためのPydanticモデル定義
"""

from pydantic import BaseModel, Field
from typing import Literal, Optional, List


class HeaderInfo(BaseModel):
    """ヘッダー情報の詳細"""
    start_row: int = Field(..., description="ヘッダーが始まる1ベースの行インデックス。")
    end_row: int = Field(..., description="ヘッダーが終わる1ベースの行インデックス。")
    header_type: Literal["single", "multi-level"] = Field(
        ..., 
        description="ヘッダーのタイプ: 'single' は単一行、'multi-level' は複数行。"
    )


class SheetAnalysisResult(BaseModel):
    """各シートの分析結果"""
    sheet_name: str = Field(..., description="分析対象のExcelシート名。")
    
    sheet_type: Literal["table", "form", "mixed", "unknown"] = Field(
        ..., 
        description="シートの主要な目的の分類。"
    )
    
    header_info: Optional[HeaderInfo] = Field(
        None, 
        description="検出されたテーブルヘッダーの詳細。テーブルでない場合やヘッダーが見つからない場合はnull。"
    )
    
    reasoning: str = Field(..., description="分類とヘッダー検出に関する簡潔な説明。")


class ExcelAnalysisOutput(BaseModel):
    """Excelファイル全体の分析結果"""
    sheets: List[SheetAnalysisResult] = Field(
        ..., 
        description="Excelファイル内の各シートの分析結果。"
    )

