"""
数据转换器
将不同格式的Excel数据转换为统一的PCRDataModel格式
"""
import pandas as pd
import numpy as np
from typing import Dict, List, Optional
import re

from data_model import PCRDataModel, WellData


class DataConverter:
    """数据转换器基类"""
    
    def convert(self, parsed_data: Dict) -> PCRDataModel:
        """将解析的数据转换为统一格式"""
        raise NotImplementedError


class VendorAConverter(DataConverter):
    """厂商A格式转换器"""
    
    def convert(self, parsed_data: Dict) -> PCRDataModel:
        """转换厂商A格式数据"""
        model = PCRDataModel()
        model.experiment_info = parsed_data.get('experiment_info', {})
        model.plate_type = "96"
        
        # 从扩增曲线数据中提取
        if 'amplification_data' in parsed_data:
            df = parsed_data['amplification_data']
            if not df.empty:
                # 检查数据格式：如果是已经包含Well和Channel列的格式
                if 'Well' in df.columns and 'Channel' in df.columns:
                    # 按孔位和通道分组
                    for (well_name, channel_name), group_df in df.groupby(['Well', 'Channel']):
                        if pd.isna(well_name) or pd.isna(channel_name):
                            continue
                        
                        well_name = str(well_name).strip()
                        channel_name = str(channel_name).strip()
                        
                        # 跳过列名（不应该作为通道名）
                        if channel_name in ['Well', 'Channel', 'Amplification', 'Value', 'Cycle', 'RawValue']:
                            continue
                        
                        # 获取或创建孔位数据
                        if well_name not in model.wells:
                            well = WellData(well_name=well_name)
                            model.add_well(well)
                        else:
                            well = model.get_well(well_name)
                        
                        # 添加通道数据
                        if 'Cycle' in group_df.columns:
                            # 按Cycle排序
                            group_df = group_df.sort_values('Cycle')
                            cycles = group_df['Cycle'].tolist()
                            # 只设置一次循环数（所有通道共享）
                            if not well.cycles:
                                well.cycles = cycles
                        
                        # 获取扩增值
                        if 'Amplification' in group_df.columns:
                            values = group_df['Amplification'].tolist()
                        elif 'Value' in group_df.columns:
                            values = group_df['Value'].tolist()
                        else:
                            continue
                        
                        # 过滤NaN并确保长度正确（42个循环）
                        values = [v if pd.notna(v) else 0.0 for v in values]
                        if well.cycles and len(values) != len(well.cycles):
                            if len(values) > len(well.cycles):
                                values = values[:len(well.cycles)]
                            else:
                                values.extend([0.0] * (len(well.cycles) - len(values)))
                        
                        well.channels[channel_name] = values
        
        # 处理原始数据
        if 'raw_data' in parsed_data:
            df_raw = parsed_data['raw_data']
            if not df_raw.empty and 'Well' in df_raw.columns and 'Channel' in df_raw.columns:
                for (well_name, channel_name), group_df in df_raw.groupby(['Well', 'Channel']):
                    if pd.isna(well_name) or pd.isna(channel_name):
                        continue
                    
                    well_name = str(well_name).strip()
                    channel_name = str(channel_name).strip()
                    
                    # 跳过列名（不应该作为通道名）
                    if channel_name in ['Well', 'Channel', 'Amplification', 'Value', 'Cycle', 'RawValue']:
                        continue
                    
                    # 获取或创建孔位数据
                    if well_name not in model.wells:
                        well = WellData(well_name=well_name)
                        model.add_well(well)
                    else:
                        well = model.get_well(well_name)
                    
                    # 添加原始数据
                    if 'RawValue' in group_df.columns:
                        # 按Cycle排序
                        group_df = group_df.sort_values('Cycle')
                        values = group_df['RawValue'].tolist()
                        values = [v if pd.notna(v) else 0.0 for v in values]
                        
                        # 确保数据长度正确（42个循环）
                        if well.cycles:
                            if len(values) > len(well.cycles):
                                values = values[:len(well.cycles)]
                            elif len(values) < len(well.cycles):
                                values.extend([0.0] * (len(well.cycles) - len(values)))
                        
                        well.raw_channels[channel_name] = values
                    
                    # 设置循环数（如果还没有）
                    if not well.cycles and 'Cycle' in group_df.columns:
                        well.cycles = sorted(group_df['Cycle'].unique().tolist())
                        
                        # 添加Ct值（如果有）
                        well_data_map = parsed_data.get('well_data', {})
                        if well_name in well_data_map and 'ct' in well_data_map[well_name]:
                            well.ct_values[channel_name] = well_data_map[well_name]['ct']
                
                else:
                    # 原始数据的旧格式处理（不应该执行到这里，因为已经有Well和Channel列）
                    # 如果原始数据没有Well和Channel列，跳过处理
                    pass
        
        return model


class DefaultConverter(DataConverter):
    """默认格式转换器"""
    
    def convert(self, parsed_data: Dict) -> PCRDataModel:
        """转换默认格式数据"""
        model = PCRDataModel()
        model.experiment_info = parsed_data.get('experiment_info', {})
        model.plate_type = "96"
        
        # 从扩增数据中提取
        if 'amplification_data' in parsed_data:
            df = parsed_data['amplification_data']
            if not df.empty:
                channels = [col for col in df.columns if col != 'Cycle']
                
                if 'Cycle' in df.columns:
                    cycles = df['Cycle'].tolist()
                else:
                    cycles = list(range(1, len(df) + 1))
                
                # 检查是否有孔位列
                well_col = None
                for col in ['Well', '孔位', 'well', 'WellName']:
                    if col in df.columns:
                        well_col = col
                        break
                
                if well_col:
                    # 按孔位分组
                    for well_name in df[well_col].unique():
                        if pd.isna(well_name):
                            continue
                        
                        well_df = df[df[well_col] == well_name]
                        well = WellData(well_name=str(well_name), cycles=cycles)
                        
                        for channel in channels:
                            if channel in well_df.columns:
                                values = well_df[channel].values.tolist()
                                values = [v if pd.notna(v) else 0.0 for v in values]
                                well.channels[channel] = values
                        
                        model.add_well(well)
                else:
                    # 没有孔位信息，创建默认孔位
                    well_name = "A1"
                    well = WellData(well_name=well_name, cycles=cycles)
                    
                    for channel in channels:
                        if channel in df.columns:
                            values = df[channel].values.tolist()
                            values = [v if pd.notna(v) else 0.0 for v in values]
                            well.channels[channel] = values
                    
                    model.add_well(well)
        
        return model


class ConverterFactory:
    """转换器工厂"""
    
    @staticmethod
    def get_converter(vendor_type: str) -> DataConverter:
        """根据厂商类型获取对应的转换器"""
        converters = {
            'vendor_a': VendorAConverter(),
            'default': DefaultConverter(),
        }
        return converters.get(vendor_type, DefaultConverter())
    
    @staticmethod
    def convert_data(parsed_data: Dict, vendor_type: str = 'default') -> PCRDataModel:
        """转换数据"""
        converter = ConverterFactory.get_converter(vendor_type)
        return converter.convert(parsed_data)

