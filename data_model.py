"""
PCR数据统一格式模型
设计为通用格式，方便不同场景的数据转换
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional
import pandas as pd
import numpy as np


@dataclass
class WellData:
    """单个孔位的数据"""
    well_name: str  # 孔位名称，如 "A1", "B2"
    channels: Dict[str, List[float]] = field(default_factory=dict)  # 通道数据 {channel_name: [values]} - 扩增数据
    raw_channels: Dict[str, List[float]] = field(default_factory=dict)  # 原始通道数据 {channel_name: [values]} - 原始数据
    cycles: List[int] = field(default_factory=list)  # 循环数列表
    ct_values: Dict[str, float] = field(default_factory=dict)  # Ct值 {channel_name: ct_value}
    metadata: Dict = field(default_factory=dict)  # 其他元数据
    
    def get_channel_data(self, channel_name: str) -> Optional[List[float]]:
        """获取指定通道的数据"""
        return self.channels.get(channel_name)
    
    def has_channel(self, channel_name: str) -> bool:
        """检查是否有指定通道的数据"""
        return channel_name in self.channels and len(self.channels[channel_name]) > 0


@dataclass
class PCRDataModel:
    """PCR数据统一模型"""
    wells: Dict[str, WellData] = field(default_factory=dict)  # {well_name: WellData}
    experiment_info: Dict = field(default_factory=dict)  # 实验信息
    plate_type: str = "96"  # 孔板类型：96或384
    
    def add_well(self, well_data: WellData):
        """添加孔位数据"""
        self.wells[well_data.well_name] = well_data
    
    def get_well(self, well_name: str) -> Optional[WellData]:
        """获取指定孔位的数据"""
        return self.wells.get(well_name)
    
    def get_all_channels(self) -> List[str]:
        """获取所有通道名称"""
        channels = set()
        for well in self.wells.values():
            # 只添加实际的通道名，排除数据列名
            for ch in well.channels.keys():
                if ch not in ['Well', 'Channel', 'Amplification', 'Value', 'Cycle']:
                    channels.add(ch)
            for ch in well.raw_channels.keys():
                if ch not in ['Well', 'Channel', 'RawValue', 'Value', 'Cycle']:
                    channels.add(ch)
        return sorted(list(channels))
    
    def get_wells_by_channels(self, channel_names: List[str]) -> Dict[str, WellData]:
        """获取包含指定通道的孔位"""
        result = {}
        for well_name, well in self.wells.items():
            if any(well.has_channel(ch) for ch in channel_names):
                result[well_name] = well
        return result
    
    def to_dataframe(self, well_names: Optional[List[str]] = None, 
                    channel_names: Optional[List[str]] = None) -> pd.DataFrame:
        """
        转换为DataFrame格式，用于绘图
        返回格式：Cycle, Well, Channel, Value
        """
        rows = []
        
        # 确定要处理的孔位
        target_wells = well_names if well_names else list(self.wells.keys())
        
        # 确定要处理的通道
        all_channels = self.get_all_channels()
        target_channels = channel_names if channel_names else all_channels
        
        for well_name in target_wells:
            well = self.get_well(well_name)
            if not well:
                continue
            
            # 获取循环数
            cycles = well.cycles if well.cycles else list(range(1, 41))  # 默认40个循环
            
            for channel_name in target_channels:
                # 跳过列名
                if channel_name in ['Well', 'Channel', 'Amplification', 'Value', 'Cycle', 'RawValue']:
                    continue
                
                # 处理HEX和VIC的等价关系
                actual_channel = channel_name
                if not well.has_channel(channel_name):
                    # 如果选择HEX但找不到HEX数据，尝试查找VIC数据
                    if channel_name == 'HEX' and well.has_channel('VIC'):
                        actual_channel = 'VIC'
                    # 如果选择VIC但找不到VIC数据，尝试查找HEX数据
                    elif channel_name == 'VIC' and well.has_channel('HEX'):
                        actual_channel = 'HEX'
                    else:
                        continue
                
                values = well.get_channel_data(actual_channel)
                if not values:
                    continue
                
                # 确保循环数和数据长度一致
                min_len = min(len(cycles), len(values))
                for i in range(min_len):
                    rows.append({
                        'Cycle': cycles[i],
                        'Well': well_name,
                        'Channel': channel_name,  # 使用用户选择的通道名，而不是actual_channel
                        'Value': values[i]
                    })
        
        return pd.DataFrame(rows)
    
    def get_amplification_data(self, well_names: Optional[List[str]] = None,
                              channel_names: Optional[List[str]] = None) -> pd.DataFrame:
        """
        获取扩增曲线数据（经过处理的荧光值）
        返回格式：Cycle, Well, Channel, Amplification
        """
        df = self.to_dataframe(well_names, channel_names)
        if df.empty:
            return df
        
        # 扩增曲线通常是原始荧光值的对数或归一化处理
        # 这里可以根据实际需求进行数据处理
        df['Amplification'] = df['Value']
        
        return df
    
    def get_raw_data(self, well_names: Optional[List[str]] = None,
                    channel_names: Optional[List[str]] = None) -> pd.DataFrame:
        """
        获取原始曲线数据（未处理的原始荧光值）
        返回格式：Cycle, Well, Channel, RawValue
        """
        rows = []
        
        # 确定要处理的孔位
        target_wells = well_names if well_names else list(self.wells.keys())
        
        # 确定要处理的通道
        all_channels = self.get_all_channels()
        target_channels = channel_names if channel_names else all_channels
        
        for well_name in target_wells:
            well = self.get_well(well_name)
            if not well:
                continue
            
            # 获取循环数
            cycles = well.cycles if well.cycles else list(range(1, 43))  # 默认42个循环
            
            for channel_name in target_channels:
                # 处理HEX和VIC的等价关系
                actual_channel = channel_name
                if channel_name not in well.raw_channels and channel_name not in well.channels:
                    # 如果选择HEX但找不到HEX数据，尝试查找VIC数据
                    if channel_name == 'HEX':
                        if 'VIC' in well.raw_channels:
                            actual_channel = 'VIC'
                        elif 'VIC' in well.channels:
                            actual_channel = 'VIC'
                        else:
                            continue
                    # 如果选择VIC但找不到VIC数据，尝试查找HEX数据
                    elif channel_name == 'VIC':
                        if 'HEX' in well.raw_channels:
                            actual_channel = 'HEX'
                        elif 'HEX' in well.channels:
                            actual_channel = 'HEX'
                        else:
                            continue
                    else:
                        continue
                
                # 优先使用raw_channels，如果没有则使用channels
                if actual_channel in well.raw_channels:
                    values = well.raw_channels[actual_channel]
                elif actual_channel in well.channels:
                    values = well.channels[actual_channel]
                else:
                    continue
                
                if not values:
                    continue
                
                # 确保循环数和数据长度一致
                min_len = min(len(cycles), len(values))
                for i in range(min_len):
                    rows.append({
                        'Cycle': cycles[i],
                        'Well': well_name,
                        'Channel': channel_name,  # 使用用户选择的通道名，而不是actual_channel
                        'RawValue': values[i]
                    })
        
        return pd.DataFrame(rows)

