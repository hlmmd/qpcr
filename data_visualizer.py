"""
PCR数据可视化模块
"""
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from matplotlib.figure import Figure
from typing import List
from data_model import PCRDataModel

# 配置matplotlib支持中文
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题


class DataVisualizer:
    """数据可视化类"""
    
    def plot_curves(self, figure, data_model: PCRDataModel, 
                   well_names: List[str], channel_names: List[str],
                   curve_type: str = 'amplification'):
        """
        绘制曲线（扩增曲线或原始曲线）
        
        Args:
            figure: matplotlib figure对象
            data_model: PCR数据模型
            well_names: 要显示的孔位列表
            channel_names: 要显示的通道列表
            curve_type: 曲线类型 'amplification' 或 'raw'
        """
        figure.clear()
        
        if not data_model.wells:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, '无数据可显示', 
                   ha='center', va='center', fontsize=14)
            ax.set_xticks([])
            ax.set_yticks([])
            return
        
        ax = figure.add_subplot(111)
        
        # 根据曲线类型获取数据
        if curve_type == 'amplification':
            df = data_model.get_amplification_data(well_names, channel_names)
            y_column = 'Amplification'
            y_label = '荧光值'
            title = 'PCR扩增曲线'
        else:  # raw
            df = data_model.get_raw_data(well_names, channel_names)
            y_column = 'RawValue'
            y_label = '原始荧光值 (Raw Fluorescence)'
            title = 'PCR原始曲线'
        
        if df.empty:
            ax.text(0.5, 0.5, '无数据可显示', 
                   ha='center', va='center', fontsize=14)
            ax.set_xticks([])
            ax.set_yticks([])
            return
        
        # 颜色映射
        channel_colors = {
            'HEX': '#1f77b4',
            'CY5': '#ff7f0e',
            'ROX': '#2ca02c',
            'FAM': '#d62728',
            'VIC': '#9467bd',
            'CY3': '#8c564b'
        }
        
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b']
        
        # 为每个通道和孔位组合绘制曲线
        plotted_count = 0  # 用于颜色索引
        
        for channel in channel_names:
            channel_df = df[df['Channel'] == channel]
            if channel_df.empty:
                continue
            
            # 获取颜色
            color = channel_colors.get(channel, colors[plotted_count % len(colors)])
            
            # 如果显示多个孔位，为每个孔位绘制一条线
            if len(well_names) > 1:
                for well_name in well_names:
                    well_df = channel_df[channel_df['Well'] == well_name]
                    if well_df.empty:
                        continue
                    
                    # 按Cycle排序
                    well_df = well_df.sort_values('Cycle')
                    
                    ax.plot(well_df['Cycle'], well_df[y_column],
                           color=color,
                           linewidth=2, alpha=0.7)
                    plotted_count += 1
            else:
                # 只显示一个孔位或所有数据合并
                if len(well_names) == 1:
                    well_df = channel_df[channel_df['Well'] == well_names[0]]
                    if not well_df.empty:
                        well_df = well_df.sort_values('Cycle')
                    else:
                        continue
                else:
                    # 合并所有孔位的数据（取平均值）
                    if 'Well' in channel_df.columns:
                        well_df = channel_df.groupby('Cycle')[y_column].mean().reset_index()
                        # 获取Cycle列
                        cycles = channel_df.groupby('Cycle')['Cycle'].first().reset_index()
                        well_df['Cycle'] = cycles['Cycle'].values
                    else:
                        well_df = channel_df.sort_values('Cycle')
                
                if not well_df.empty:
                    ax.plot(well_df['Cycle'], well_df[y_column],
                           color=color,
                           linewidth=2)
                    plotted_count += 1
        
        ax.set_xlabel('循环数', fontsize=12)
        ax.set_ylabel(y_label, fontsize=12)
        
        # 设置y轴范围
        if curve_type == 'amplification':
            ax.set_ylim(0, 5000)  # 扩增曲线：0-5000
        else:  # raw
            ax.set_ylim(0, 10000)  # 原始曲线：0-10000
        
        # 添加孔位信息到标题
        if well_names:
            if len(well_names) == 1:
                title += f" - {well_names[0]}"
            elif len(well_names) <= 3:
                title += f" - {', '.join(well_names)}"
            else:
                title += f" - {len(well_names)}个孔位"
        
        ax.set_title(title, fontsize=14, fontweight='bold')
        
        # 不显示图例
        
        ax.grid(True, alpha=0.3)
        
        figure.tight_layout()
    
    def plot_amplification_curves(self, figure, data_model: PCRDataModel,
                                 well_names: List[str], channel_names: List[str]):
        """绘制扩增曲线"""
        self.plot_curves(figure, data_model, well_names, channel_names, 'amplification')
    
    def plot_raw_curves(self, figure, data_model: PCRDataModel,
                       well_names: List[str], channel_names: List[str]):
        """绘制原始曲线"""
        self.plot_curves(figure, data_model, well_names, channel_names, 'raw')
    
    def plot_amplification_curves_old(self, figure, data_df):
        """绘制扩增曲线（旧版本，保持兼容）"""
        figure.clear()
        
        if data_df.empty:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, '无数据可显示', 
                   ha='center', va='center', fontsize=14)
            ax.set_xticks([])
            ax.set_yticks([])
            return
        
        # 创建子图
        ax = figure.add_subplot(111)
        
        # 获取通道列（排除Cycle列）
        channels = [col for col in data_df.columns if col != 'Cycle']
        
        if 'Cycle' not in data_df.columns:
            # 如果没有Cycle列，使用索引
            cycles = range(1, len(data_df) + 1)
        else:
            cycles = data_df['Cycle'].values
        
        # 为每个通道绘制曲线
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b']
        
        for i, channel in enumerate(channels):
            if channel in data_df.columns:
                values = data_df[channel].values
                # 过滤NaN值
                valid_mask = ~np.isnan(values)
                if np.any(valid_mask):
                    ax.plot(cycles[valid_mask], values[valid_mask], 
                           label=channel, 
                           color=colors[i % len(colors)], 
                           linewidth=2)
        
        ax.set_xlabel('循环数 (Cycle)', fontsize=12)
        ax.set_ylabel('荧光值 / Ct值', fontsize=12)
        ax.set_title('PCR扩增曲线', fontsize=14, fontweight='bold')
        ax.legend(loc='best', fontsize=10)
        ax.grid(True, alpha=0.3)
        
        figure.tight_layout()
    
    def plot_ct_values(self, figure, data_df):
        """绘制Ct值柱状图"""
        figure.clear()
        
        if data_df.empty:
            return
        
        ax = figure.add_subplot(111)
        
        # 获取通道列
        channels = [col for col in data_df.columns if col != 'Cycle']
        
        # 计算每个通道的平均Ct值（排除NaN）
        ct_values = []
        channel_names = []
        
        for channel in channels:
            values = data_df[channel].values
            valid_values = values[~np.isnan(values)]
            if len(valid_values) > 0:
                # 如果是Ct值，取第一个有效值；如果是荧光值，可能需要计算
                ct_values.append(valid_values[0] if len(valid_values) > 0 else np.nan)
                channel_names.append(channel)
        
        if ct_values:
            bars = ax.bar(channel_names, ct_values, color=['#1f77b4', '#ff7f0e', '#2ca02c'])
            ax.set_ylabel('Ct值', fontsize=12)
            ax.set_title('各通道Ct值', fontsize=14, fontweight='bold')
            ax.grid(True, alpha=0.3, axis='y')
            
            # 添加数值标签
            for bar in bars:
                height = bar.get_height()
                if not np.isnan(height):
                    ax.text(bar.get_x() + bar.get_width()/2., height,
                           f'{height:.2f}',
                           ha='center', va='bottom', fontsize=10)
        
        figure.tight_layout()
