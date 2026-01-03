"""
Excel文件解析器
支持解析不同厂商的PCR数据格式
"""
import pandas as pd
import numpy as np
import openpyxl
from pathlib import Path
import re


class ExcelParser:
    """Excel文件解析器基类"""
    
    def __init__(self):
        self.parsers = {
            'default': DefaultParser(),
            'vendor_a': VendorAParser(),
        }
        
        # 尝试导入VendorBParser（如果存在）
        try:
            from VendorBParser import VendorBParser
            self.parsers['vendor_b'] = VendorBParser()
        except ImportError:
            pass
    
    def parse(self, file_path):
        """解析Excel文件，自动识别格式"""
        # 尝试识别厂商格式
        vendor = self.detect_vendor(file_path)
        
        # 使用对应的解析器
        parser = self.parsers.get(vendor, self.parsers['default'])
        return parser.parse(file_path)
    
    def detect_vendor(self, file_path):
        """检测Excel文件的厂商格式"""
        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.sheetnames
        
        # 根据工作表名称和内容特征识别
        if any('实验数据' in name or '扩增曲线' in name for name in sheet_names):
            return 'vendor_a'  # 示例：基于中文工作表名
        
        # 可以添加更多识别逻辑
        return 'default'


class BaseParser:
    """解析器基类"""
    
    def parse(self, file_path):
        """解析文件，返回结构化数据"""
        raise NotImplementedError
    
    def extract_experiment_info(self, df):
        """提取实验信息"""
        info = {}
        # 通用提取逻辑
        return info
    
    def extract_amplification_data(self, df):
        """提取扩增数据"""
        raise NotImplementedError


class DefaultParser(BaseParser):
    """默认解析器"""
    
    def parse(self, file_path):
        """解析标准格式的Excel文件"""
        wb = openpyxl.load_workbook(file_path)
        result = {
            'sheets': {},
            'experiment_info': {},
            'amplification_data': pd.DataFrame()
        }
        
        for sheet_name in wb.sheetnames:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            result['sheets'][sheet_name] = df
            
            # 尝试提取实验信息
            info = self.extract_experiment_info(df)
            result['experiment_info'].update(info)
            
            # 尝试提取扩增数据
            amp_data = self.extract_amplification_data(df)
            if not amp_data.empty:
                result['amplification_data'] = amp_data
        
        return result
    
    def extract_experiment_info(self, df):
        """提取实验信息"""
        info = {}
        
        # 查找关键信息行
        for idx, row in df.iterrows():
            row_str = ' '.join([str(x) for x in row if pd.notna(x)])
            
            # 提取开始时间
            if '开始时间' in row_str or '起始时间' in row_str:
                for i, val in enumerate(row):
                    if pd.notna(val) and i < len(row) - 1:
                        next_val = row[i + 1]
                        if pd.notna(next_val):
                            info['开始时间'] = str(next_val)
                            break
            
            # 提取结束时间
            if '结束时间' in row_str or '完成时间' in row_str:
                for i, val in enumerate(row):
                    if pd.notna(val) and i < len(row) - 1:
                        next_val = row[i + 1]
                        if pd.notna(next_val):
                            info['结束时间'] = str(next_val)
                            break
        
        return info
    
    def extract_amplification_data(self, df):
        """提取扩增数据"""
        # 查找数据区域
        data_start_row = None
        channels = []
        
        # 查找通道名称行（如HEX, CY5, ROX等）
        for idx, row in df.iterrows():
            row_values = [str(x).upper() for x in row if pd.notna(x)]
            
            # 检查是否包含常见通道名
            common_channels = ['HEX', 'CY5', 'ROX', 'FAM', 'VIC', 'CY3']
            found_channels = [ch for ch in common_channels if any(ch in val for val in row_values)]
            
            if found_channels:
                data_start_row = idx
                # 提取通道信息
                for i, val in enumerate(row):
                    if pd.notna(val):
                        val_str = str(val).upper()
                        for ch in common_channels:
                            if ch in val_str:
                                channels.append((i, ch))
                break
        
        if data_start_row is None:
            return pd.DataFrame()
        
        # 提取数据
        data_rows = []
        cycle_col = None
        
        # 查找循环数列
        for idx in range(data_start_row, min(data_start_row + 50, len(df))):
            row = df.iloc[idx]
            first_val = row.iloc[0] if len(row) > 0 else None
            
            # 检查是否是数字（循环数）
            if pd.notna(first_val):
                try:
                    cycle_num = float(first_val)
                    if 1 <= cycle_num <= 50:  # 合理的循环数范围
                        cycle_col = 0
                        break
                except:
                    pass
        
        # 提取数据行
        for idx in range(data_start_row + 1, len(df)):
            row = df.iloc[idx]
            row_data = {}
            
            # 提取循环数
            if cycle_col is not None and pd.notna(row.iloc[cycle_col]):
                try:
                    row_data['Cycle'] = int(float(row.iloc[cycle_col]))
                except:
                    continue
            
            # 提取各通道数据
            for col_idx, channel_name in channels:
                if col_idx < len(row):
                    val = row.iloc[col_idx]
                    if pd.notna(val):
                        val_str = str(val).upper()
                        # 处理NoCt等特殊值
                        if 'NOCT' in val_str or 'N/A' in val_str:
                            row_data[channel_name] = np.nan
                        else:
                            try:
                                row_data[channel_name] = float(val)
                            except:
                                row_data[channel_name] = np.nan
            
            if row_data:
                data_rows.append(row_data)
        
        if data_rows:
            return pd.DataFrame(data_rows)
        return pd.DataFrame()


class VendorAParser(BaseParser):
    """厂商A格式解析器（基于示例文件格式）"""
    
    def parse(self, file_path):
        """解析厂商A的Excel格式"""
        result = {
            'sheets': {},
            'experiment_info': {},
            'amplification_data': pd.DataFrame()
        }
        
        wb = openpyxl.load_workbook(file_path)
        
        # 解析"实验数据"工作表
        if '实验数据' in wb.sheetnames:
            df_exp = pd.read_excel(file_path, sheet_name='实验数据', header=None)
            result['sheets']['实验数据'] = df_exp
            result['experiment_info'] = self.extract_experiment_info(df_exp)
            # 提取孔位数据
            result['well_data'] = self.extract_well_data(df_exp)
            # 从实验数据工作表中提取扩增数据
            result['amplification_data'] = self.extract_amplification_data_from_exp(df_exp)
            # 从实验数据工作表中提取原始数据
            result['raw_data'] = self.extract_raw_data_from_exp(df_exp)
        
        # 解析"扩增曲线"工作表（如果存在）
        if '扩增曲线' in wb.sheetnames:
            df_curve = pd.read_excel(file_path, sheet_name='扩增曲线', header=None)
            result['sheets']['扩增曲线'] = df_curve
            # 如果实验数据中没有扩增数据，则从扩增曲线工作表提取
            if result['amplification_data'].empty:
                result['amplification_data'] = self.extract_amplification_data(df_curve)
        
        return result
    
    def extract_experiment_info(self, df):
        """提取实验信息"""
        info = {}
        
        # 查找关键信息
        for idx in range(min(20, len(df))):
            row = df.iloc[idx]
            first_col = row.iloc[0] if len(row) > 0 else None
            
            if pd.notna(first_col):
                first_str = str(first_col)
                
                # 开始时间
                if '开始时间' in first_str or '起始时间' in first_str:
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        info['开始时间'] = str(row.iloc[1])
                
                # 结束时间
                if '结束时间' in first_str or '完成时间' in first_str:
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        info['结束时间'] = str(row.iloc[1])
                
                # 实验名称
                if '实验名称' in first_str:
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        info['实验名称'] = str(row.iloc[1])
        
        return info
    
    def extract_amplification_data(self, df):
        """从扩增曲线工作表提取数据"""
        # 查找通道行（包含HEX, CY5, ROX等）
        channel_row_idx = None
        channels = []
        
        # 扩大搜索范围
        for idx in range(min(30, len(df))):
            row = df.iloc[idx]
            row_str = ' '.join([str(x).upper() for x in row if pd.notna(x)])
            
            # 检查是否包含通道名
            if any(ch in row_str for ch in ['HEX', 'CY5', 'ROX', 'FAM']):
                channel_row_idx = idx
                # 记录通道位置
                for col_idx, val in enumerate(row):
                    if pd.notna(val):
                        val_str = str(val).upper()
                        if 'HEX' in val_str:
                            channels.append((col_idx, 'HEX'))
                        elif 'CY5' in val_str:
                            channels.append((col_idx, 'CY5'))
                        elif 'ROX' in val_str:
                            channels.append((col_idx, 'ROX'))
                        elif 'FAM' in val_str:
                            channels.append((col_idx, 'FAM'))
                break
        
        # 如果没找到通道行，尝试查找数据区域
        if channel_row_idx is None:
            # 尝试查找包含数字数据的行
            for idx in range(min(30, len(df))):
                row = df.iloc[idx]
                # 检查是否包含数字（可能是数据行）
                numeric_count = 0
                for val in row:
                    if pd.notna(val):
                        try:
                            float(val)
                            numeric_count += 1
                        except:
                            pass
                # 如果一行中有多个数字，可能是数据行
                if numeric_count >= 3:
                    # 假设第一列是循环数，其他列是通道数据
                    channel_row_idx = idx - 1  # 假设上一行是通道名
                    # 尝试从列索引推断通道
                    for col_idx in range(1, min(10, len(row))):
                        # 根据列位置分配通道名（如果找不到通道名）
                        if col_idx == 1:
                            channels.append((col_idx, 'HEX'))
                        elif col_idx == 2:
                            channels.append((col_idx, 'CY5'))
                        elif col_idx == 3:
                            channels.append((col_idx, 'ROX'))
                        elif col_idx == 4:
                            channels.append((col_idx, 'FAM'))
                    break
        
        if channel_row_idx is None or not channels:
            print(f"警告: 未找到通道信息，尝试默认解析...")
            # 如果还是找不到，尝试从数据中推断
            # 查找第一个包含数字的行
            for idx in range(len(df)):
                row = df.iloc[idx]
                first_val = row.iloc[0] if len(row) > 0 else None
                if pd.notna(first_val):
                    try:
                        cycle_num = float(first_val)
                        if 1 <= cycle_num <= 50:  # 合理的循环数
                            channel_row_idx = idx
                            # 假设后续列是通道数据
                            for col_idx in range(1, min(7, len(row))):
                                channels.append((col_idx, ['HEX', 'CY5', 'ROX', 'FAM', 'VIC', 'CY3'][col_idx-1]))
                            break
                    except:
                        pass
        
        if channel_row_idx is None or not channels:
            print(f"错误: 无法找到通道信息")
            return pd.DataFrame()
        
        print(f"找到通道行: {channel_row_idx}, 通道: {channels}")
        
        # 提取数据
        data_rows = []
        well_data_map = {}  # 存储每个孔位的数据 {well_name: {cycle: {channel: value}}}
        
        # 从通道行之后开始提取数据
        for idx in range(channel_row_idx + 1, len(df)):
            row = df.iloc[idx]
            row_data = {}
            
            # 尝试提取循环数（通常在某一列）
            cycle_num = idx - channel_row_idx
            if cycle_num > 0:
                row_data['Cycle'] = cycle_num
            
            # 提取各通道的Ct值或荧光值
            for col_idx, channel_name in channels:
                if col_idx < len(row):
                    val = row.iloc[col_idx]
                    if pd.notna(val):
                        val_str = str(val).upper()
                        if 'NOCT' in val_str or 'N/A' in val_str or val_str == 'NAN':
                            row_data[channel_name] = np.nan
                        else:
                            try:
                                row_data[channel_name] = float(val)
                            except:
                                pass
            
            if row_data and 'Cycle' in row_data:
                data_rows.append(row_data)
        
        if data_rows:
            result_df = pd.DataFrame(data_rows)
            # 尝试从实验数据工作表中提取孔位信息
            return result_df
        return pd.DataFrame()
    
    def extract_well_data(self, df_exp):
        """从实验数据工作表中提取孔位数据"""
        well_data = {}
        
        # 查找包含孔位信息的区域
        # 通常孔位信息在表格的某个区域
        for idx in range(len(df_exp)):
            row = df_exp.iloc[idx]
            # 查找包含孔位标识的行（如A1, B2等）
            for col_idx, val in enumerate(row):
                if pd.notna(val):
                    val_str = str(val).strip()
                    # 检查是否是孔位格式（如A1, B12等）
                    if re.match(r'^[A-H][0-9]{1,2}$', val_str, re.IGNORECASE):
                        well_name = val_str.upper()
                        # 尝试提取该孔位的Ct值等信息
                        well_info = {'well': well_name}
                        
                        # 在同一行或相邻行查找Ct值
                        for next_col in range(col_idx + 1, min(col_idx + 10, len(row))):
                            next_val = row.iloc[next_col] if next_col < len(row) else None
                            if pd.notna(next_val):
                                try:
                                    ct_val = float(next_val)
                                    if 0 < ct_val < 50:  # 合理的Ct值范围
                                        well_info['ct'] = ct_val
                                        break
                                except:
                                    pass
                        
                        well_data[well_name] = well_info
        
        return well_data
    
    def extract_amplification_data_from_exp(self, df):
        """从实验数据工作表中提取扩增数据"""
        # 查找表头行（包含"反应孔"、"通道"、"Ct"等）
        header_row_idx = None
        well_col_idx = None
        channel_col_idx = None
        ct_col_idx = None
        data_start_col = None
        
        # 查找表头行（通常是第13行，索引13）
        for idx in range(min(20, len(df))):
            row = df.iloc[idx]
            row_str = ' '.join([str(x) for x in row if pd.notna(x)])
            
            # 查找包含"反应孔"的行（通常是第13行）
            if '反应孔' in row_str:
                header_row_idx = idx
                
                # 直接设置已知的列索引（基于实际数据格式）
                well_col_idx = 0  # 第一列是反应孔
                channel_col_idx = 6  # 第7列是染色（FAM等），第6列是通道编号
                ct_col_idx = 12  # 第13列是Ct
                
                # 查找数据开始列（查找包含"1.0"的列，通常是第39列）
                for col_idx in range(35, min(45, len(row))):
                    if pd.notna(row.iloc[col_idx]):
                        val_str = str(row.iloc[col_idx])
                        if val_str == '1.0' or val_str == '1.00':
                            data_start_col = col_idx
                            break
                
                # 如果还是没找到，使用默认值39
                if data_start_col is None:
                    # 检查下一行是否有数据
                    if idx + 1 < len(df):
                        next_row = df.iloc[idx + 1]
                        for col_idx in range(35, min(45, len(next_row))):
                            val = next_row.iloc[col_idx]
                            if pd.notna(val):
                                try:
                                    num_val = float(val)
                                    if -100 < num_val < 10000:
                                        data_start_col = col_idx
                                        break
                                except:
                                    pass
                    if data_start_col is None:
                        data_start_col = 39  # 默认值
                
                break
        
        if header_row_idx is None:
            print("未找到表头行")
            return pd.DataFrame()
        
        print(f"找到表头行: {header_row_idx}, 孔位列: {well_col_idx}, 通道列: {channel_col_idx}, Ct列: {ct_col_idx}, 数据开始列: {data_start_col}")
        
        # 提取数据
        data_rows = []
        max_cycles = 0
        
        # 从表头行之后开始提取数据
        for idx in range(header_row_idx + 1, len(df)):
            row = df.iloc[idx]
            
            # 获取孔位
            well_name = None
            if well_col_idx is not None and well_col_idx < len(row):
                well_val = row.iloc[well_col_idx]
                if pd.notna(well_val):
                    well_str = str(well_val).strip()
                    # 检查是否是孔位格式
                    if re.match(r'^[A-H][0-9]{1,2}$', well_str, re.IGNORECASE):
                        well_name = well_str.upper()
            
            if not well_name:
                continue
            
            # 获取通道
            channel_name = None
            if channel_col_idx is not None and channel_col_idx < len(row):
                channel_val = row.iloc[channel_col_idx]
                if pd.notna(channel_val):
                    channel_name = str(channel_val).strip()
            
            # 获取Ct值
            ct_value = None
            if ct_col_idx is not None and ct_col_idx < len(row):
                ct_val = row.iloc[ct_col_idx]
                if pd.notna(ct_val):
                    try:
                        ct_value = float(ct_val)
                    except:
                        pass
            
            # 提取扩增数据（AN列到CC列，即列39到80，共42个循环）
            # AN列索引是40（pandas中索引39），CC列索引是81（pandas中索引80）
            amp_start_col = 39  # AN列
            amp_end_col = 81    # CC列（不包含，所以是81）
            
            if amp_start_col < len(row):
                cycle_num = 1
                for col_idx in range(amp_start_col, min(amp_end_col, len(row))):
                    val = row.iloc[col_idx]
                    if pd.isna(val):
                        # 如果遇到NaN，跳过
                        cycle_num += 1
                        continue
                    
                    try:
                        amp_value = float(val)
                        # 接受合理的扩增值范围（可以是负数，因为可能是ΔRn值）
                        if -100 < amp_value < 10000:
                            row_data = {
                                'Cycle': cycle_num,
                                'Well': well_name,
                                'Channel': channel_name if channel_name else 'Unknown',
                                'Amplification': amp_value
                            }
                            # 只在第一行添加CT值
                            if cycle_num == 1 and ct_value is not None:
                                row_data['Ct'] = ct_value
                            data_rows.append(row_data)
                            cycle_num += 1
                            max_cycles = max(max_cycles, cycle_num)
                    except:
                        # 如果转换失败，跳过
                        cycle_num += 1
                        continue
        
        if data_rows:
            result_df = pd.DataFrame(data_rows)
            print(f"提取到 {len(result_df)} 行扩增数据，最大循环数: {max_cycles}")
            return result_df
        
        return pd.DataFrame()
    
    def extract_raw_data_from_exp(self, df):
        """从实验数据工作表中提取原始曲线数据"""
        # 查找表头行
        header_row_idx = None
        well_col_idx = None
        channel_col_idx = None
        
        for idx in range(min(20, len(df))):
            row = df.iloc[idx]
            row_str = ' '.join([str(x) for x in row if pd.notna(x)])
            
            if '反应孔' in row_str:
                header_row_idx = idx
                well_col_idx = 0
                channel_col_idx = 6
                break
        
        if header_row_idx is None:
            return pd.DataFrame()
        
        # 提取原始数据（CE列到DT列，即列82到123，共42个循环）
        # CE列索引是83（pandas中索引82），DT列索引是124（pandas中索引123）
        raw_start_col = 82  # CE列
        raw_end_col = 124   # DT列（不包含，所以是124）
        
        data_rows = []
        
        for idx in range(header_row_idx + 1, len(df)):
            row = df.iloc[idx]
            
            # 获取孔位
            well_name = None
            if well_col_idx is not None and well_col_idx < len(row):
                well_val = row.iloc[well_col_idx]
                if pd.notna(well_val):
                    well_str = str(well_val).strip()
                    if re.match(r'^[A-H][0-9]{1,2}$', well_str, re.IGNORECASE):
                        well_name = well_str.upper()
            
            if not well_name:
                continue
            
            # 获取通道
            channel_name = None
            if channel_col_idx is not None and channel_col_idx < len(row):
                channel_val = row.iloc[channel_col_idx]
                if pd.notna(channel_val):
                    channel_name = str(channel_val).strip()
            
            # 提取原始数据
            if raw_start_col < len(row):
                cycle_num = 1
                for col_idx in range(raw_start_col, min(raw_end_col, len(row))):
                    val = row.iloc[col_idx]
                    if pd.isna(val):
                        cycle_num += 1
                        continue
                    
                    try:
                        raw_value = float(val)
                        if -100 < raw_value < 100000:
                            data_rows.append({
                                'Cycle': cycle_num,
                                'Well': well_name,
                                'Channel': channel_name if channel_name else 'Unknown',
                                'RawValue': raw_value
                            })
                            cycle_num += 1
                    except:
                        cycle_num += 1
                        continue
        
        if data_rows:
            result_df = pd.DataFrame(data_rows)
            print(f"提取到 {len(result_df)} 行原始数据")
            return result_df
        
        return pd.DataFrame()

