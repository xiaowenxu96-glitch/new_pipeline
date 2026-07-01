import pandas as pd
import openpyxl
from typing import Dict, List, Union, Tuple, Optional
import os
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DataReader:
    """
    负责从Excel文件读取数据并提供全局缓存的类，避免重复慢速 I/O
    """
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self._validate_file()
        self._sheet_cache = {}      # 缓存整表 DataFrame
        self._indicator_cache = {}  # 缓存解析后的指标字典 (sheet_name, indicator_code) -> Dict
        
    def _validate_file(self) -> None:
        """验证文件是否存在且是有效的Excel文件"""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"文件不存在: {self.file_path}")
        if not self.file_path.endswith(('.xlsx', '.xls')):
            raise ValueError(f"文件不是有效的Excel文件: {self.file_path}")
        logger.info(f"已验证源文件: {self.file_path}")
    
    def read_sheet_data(self, sheet_name: str) -> pd.DataFrame:
        """读取并缓存整个工作表的数据"""
        if sheet_name in self._sheet_cache:
            return self._sheet_cache[sheet_name]
            
        try:
            logger.info(f"开始读取源工作表 '{sheet_name}' 并存入缓存...")
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
            self._sheet_cache[sheet_name] = df
            return df
        except Exception as e:
            logger.error(f"读取工作表 '{sheet_name}' 数据时出错: {str(e)}")
            raise
            
    def read_indicator_data(self, sheet_name: str, indicator_code: str) -> Dict:
        """
        读取指定工作表中特定指标代码的数据，优先使用缓存
        """
        cache_key = (sheet_name, indicator_code)
        if cache_key in self._indicator_cache:
            return self._indicator_cache[cache_key]
            
        try:
            df = self.read_sheet_data(sheet_name)
            
            # 前10行为元数据，第11行开始为数据  
            metadata_rows = df.iloc[:10]
            data_rows = df.iloc[10:].reset_index(drop=True)
            
            # 提取元数据
            metadata = {}
            for i, row in metadata_rows.iterrows():
                metadata[row.iloc[0]] = row.iloc[1:].tolist()
            
            # 查找指标代码行
            indicator_code_row_index = None
            for i, row in metadata_rows.iterrows():
                if row.iloc[0] == '指标代码':
                    indicator_code_row_index = i
                    break
                    
            if indicator_code_row_index is None:
                raise ValueError(f"未找到指标代码行")
            
            indicator_code_row = metadata_rows.iloc[indicator_code_row_index]
            
            # 查找匹配的指标代码索引
            code_index = None
            for i, code in enumerate(indicator_code_row.iloc[1:], 1):
                if code == indicator_code:
                    code_index = i - 1
                    break
                    
            if code_index is None:
                logger.warning(f"在工作表 '{sheet_name}' 中未找到指标 '{indicator_code}'")
                return {"metadata": {}, "data": pd.DataFrame()}
                
            # 提取元数据
            indicator_metadata = {}
            for key, row in zip(metadata.keys(), metadata_rows.iterrows()):
                _, row_data = row
                if code_index + 1 < len(row_data):
                    indicator_metadata[key] = row_data.iloc[code_index + 1]
                    
            indicator_col_index = code_index + 1
            
            # 查找时间列
            time_col_index = 0
            for i in range(indicator_col_index, 0, -1):
                if indicator_code_row.iloc[i] == '指标代码':
                    time_col_index = i
                    break
                    
            # 提取该指标的数据列和对应的时间列
            if indicator_col_index < len(data_rows.columns):
                indicator_data = data_rows.iloc[:, [time_col_index, indicator_col_index]].copy()
                indicator_name = indicator_metadata.get('指标全称', f'指标_{indicator_code}')
                indicator_data.columns = ['日期', indicator_name]
                
                # 处理可能的异常日期格式
                valid_dates = pd.to_datetime(indicator_data['日期'], errors='coerce')
                indicator_data['日期'] = valid_dates
                # 过滤掉无法转换的无效日期行（如"上传下载"等脏数据）
                indicator_data = indicator_data.dropna(subset=['日期']).reset_index(drop=True)
                
                result = {
                    "metadata": indicator_metadata,
                    "data": indicator_data
                }
                # 存入缓存
                self._indicator_cache[cache_key] = result
                return result
            else:
                return {"metadata": indicator_metadata, "data": pd.DataFrame()}
                
        except Exception as e:
            logger.error(f"读取指标 '{indicator_code}' 数据时出错: {str(e)}")
            raise
