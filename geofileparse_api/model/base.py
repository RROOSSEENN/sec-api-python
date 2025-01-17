# app/models/base.py
from pydantic import BaseModel, Field
from typing import Any, Dict, Optional

class BaseCustomModel(BaseModel):
    """通用的数据模型，包含一个可接受任意结构字段的 data。"""
    data: Dict[str, Any] = Field(..., description="任意结构数据")
