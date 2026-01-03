"""
PCR结果分析软件
支持解析不同厂商的Excel格式，展示PCR扩增结果
"""
import sys
import os
from pathlib import Path
from typing import List
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QFileDialog, QTableWidget, 
                             QTableWidgetItem, QLabel, QMessageBox, QTabWidget,
                             QTextEdit, QSplitter, QGroupBox, QComboBox, QCheckBox, QRadioButton, QButtonGroup)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QColor
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib
import openpyxl

# 配置matplotlib支持中文
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

from excel_parser import ExcelParser
from data_visualizer import DataVisualizer
from plate_selector import PlateSelector
from data_model import PCRDataModel
from data_converter import ConverterFactory

# 导入项目数据
try:
    from projects_data import projects_data, channel_names as project_channel_names
except ImportError:
    projects_data = {}
    project_channel_names = ['FAM', 'VIC', 'CY5', 'ROX']


class PCRAnalyzerApp(QMainWindow):
    """PCR分析软件主窗口"""
    
    def __init__(self):
        super().__init__()
        self.current_file = None
        self.parsed_data = None
        self.data_model = None  # 统一数据模型
        self.well_data_map = {}  # 孔位数据映射 {well_name: data}
        self.selected_channels = []  # 选中的通道
        self.curve_type = 'amplification'  # 当前曲线类型
        self.selected_project = None  # 选中的项目名称
        self.judgment_results = []  # 研判结果列表
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle('PCR结果分析软件')
        self.setGeometry(100, 100, 1600, 1000)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 工具栏
        toolbar = self.create_toolbar()
        main_layout.addWidget(toolbar)
        
        # 创建标签页
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # 主分析标签页（包含孔板选择器和扩增曲线）
        self.main_tab = self.create_main_tab()
        self.tabs.addTab(self.main_tab, "PCR分析")
        
        # 数据展示标签页
        self.data_tab = self.create_data_tab()
        self.tabs.addTab(self.data_tab, "实验数据")
        
        # 状态栏
        self.statusBar().showMessage('就绪')
        
    def create_toolbar(self):
        """创建工具栏"""
        toolbar = QGroupBox()
        layout = QHBoxLayout(toolbar)
        
        # 打开文件按钮
        self.open_btn = QPushButton('打开Excel文件')
        self.open_btn.clicked.connect(self.open_file)
        layout.addWidget(self.open_btn)
        
        # 导出按钮
        self.export_btn = QPushButton('导出结果')
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        layout.addWidget(self.export_btn)
        
        layout.addStretch()
        
        # 文件信息标签
        self.file_label = QLabel('未打开文件')
        layout.addWidget(self.file_label)
        
        return toolbar
    
    def create_main_tab(self):
        """创建主分析标签页（包含孔板选择器和扩增曲线）"""
        widget = QWidget()
        # 使用垂直布局：第一行是孔板|曲线|项目，第二行是结果
        main_layout = QVBoxLayout(widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)
        
        # 第一行：孔板 | 曲线 | 项目（使用水平分割器）
        top_splitter = QSplitter(Qt.Horizontal)
        
        # 左侧：孔板选择器和控制面板
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(5, 5, 5, 5)
        
        # 孔板选择器（固定96孔板）
        self.plate_selector = PlateSelector(plate_type='96')
        self.plate_selector.well_selected.connect(self.on_well_selected)
        left_layout.addWidget(self.plate_selector)
        
        # 通道选择
        channel_group = QGroupBox("通道选择")
        channel_layout = QVBoxLayout(channel_group)
        self.channel_checkboxes = {}
        self.channel_checkboxes['HEX'] = QCheckBox('HEX')
        self.channel_checkboxes['CY5'] = QCheckBox('CY5')
        self.channel_checkboxes['ROX'] = QCheckBox('ROX')
        self.channel_checkboxes['FAM'] = QCheckBox('FAM')
        
        for checkbox in self.channel_checkboxes.values():
            checkbox.setChecked(True)
            checkbox.stateChanged.connect(self.on_channel_changed)
            channel_layout.addWidget(checkbox)
        
        left_layout.addWidget(channel_group)
        
        # 清除选择按钮
        clear_btn = QPushButton('清除选择')
        clear_btn.clicked.connect(self.clear_all_selection)
        left_layout.addWidget(clear_btn)
        
        left_layout.addStretch()
        
        # 中间：扩增曲线显示
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # 控制按钮
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("显示选项:"))
        
        # 曲线类型选择
        curve_type_group = QGroupBox("曲线类型")
        curve_type_layout = QHBoxLayout(curve_type_group)
        self.curve_type_group = QButtonGroup()
        self.amplification_radio = QRadioButton('扩增曲线')
        self.amplification_radio.setChecked(True)
        self.raw_radio = QRadioButton('原始曲线')
        self.curve_type_group.addButton(self.amplification_radio, 0)
        self.curve_type_group.addButton(self.raw_radio, 1)
        self.amplification_radio.toggled.connect(self.on_curve_type_changed)
        self.raw_radio.toggled.connect(self.on_curve_type_changed)
        curve_type_layout.addWidget(self.amplification_radio)
        curve_type_layout.addWidget(self.raw_radio)
        control_layout.addWidget(curve_type_group)
        
        self.show_all_wells_check = QCheckBox('显示所有选中孔位')
        self.show_all_wells_check.setChecked(True)
        self.show_all_wells_check.stateChanged.connect(self.update_curves)
        control_layout.addWidget(self.show_all_wells_check)
        
        control_layout.addStretch()
        right_layout.addLayout(control_layout)
        
        # 扩增曲线画布
        self.figure = Figure(figsize=(10, 7))
        self.canvas = FigureCanvas(self.figure)
        right_layout.addWidget(self.canvas)
        
        # 信息显示（已隐藏）
        self.curve_info = QTextEdit()
        self.curve_info.setReadOnly(True)
        self.curve_info.setMaximumHeight(100)
        self.curve_info.setVisible(False)  # 隐藏信息显示区域
        # right_layout.addWidget(self.curve_info)  # 不添加到布局中
        
        # 右侧：项目选择列
        project_panel = QWidget()
        project_layout = QVBoxLayout(project_panel)
        project_layout.setContentsMargins(5, 5, 5, 5)
        
        # 项目选择组
        project_group = QGroupBox("项目选择")
        project_group_layout = QVBoxLayout(project_group)
        
        # 项目复选框列表
        self.project_checkboxes = {}
        if projects_data:
            for project_name in sorted(projects_data.keys()):
                checkbox = QCheckBox(project_name)
                # 创建包装函数来传递项目名称，避免lambda闭包问题
                def make_handler(name):
                    def handler(checked):
                        self.on_project_changed(name)
                    return handler
                checkbox.stateChanged.connect(make_handler(project_name))
                self.project_checkboxes[project_name] = checkbox
                project_group_layout.addWidget(checkbox)
        else:
            no_project_label = QLabel("未找到项目数据")
            project_group_layout.addWidget(no_project_label)
        
        project_group_layout.addStretch()
        project_layout.addWidget(project_group)
        project_layout.addStretch()
        
        # 将三个面板添加到水平分割器
        top_splitter.addWidget(left_panel)
        top_splitter.addWidget(right_panel)
        top_splitter.addWidget(project_panel)
        top_splitter.setStretchFactor(0, 1)
        top_splitter.setStretchFactor(1, 2)
        top_splitter.setStretchFactor(2, 1)
        
        # 添加到主布局（第一行）
        main_layout.addWidget(top_splitter, 2)  # 占据2倍空间
        
        # 第二行：研判结果显示区域（单独一行，占据整个宽度）
        judgment_group = QGroupBox("研判结果")
        judgment_layout = QVBoxLayout(judgment_group)
        
        # 结果显示表格（列数和列标题会根据选中的项目动态设置）
        self.judgment_table = QTableWidget()
        # 初始列数会在update_judgment_results中动态设置
        # 设置最小高度，让结果表格有足够的显示空间
        self.judgment_table.setMinimumHeight(250)
        judgment_layout.addWidget(self.judgment_table)
        
        # 添加到主布局（第二行）
        main_layout.addWidget(judgment_group, 1)  # 占据1倍空间
        
        return widget
    
    def create_data_tab(self):
        """创建数据展示标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 信息显示区域
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setMaximumHeight(150)
        layout.addWidget(self.info_text)
        
        # 数据表格
        self.data_table = QTableWidget()
        layout.addWidget(self.data_table)
        
        return widget
    
    
    def clear_all_selection(self):
        """清除所有选择（孔位和通道）"""
        # 清除孔位选择
        self.plate_selector.clear_selection()
        
        # 清除通道选择（取消选中所有通道复选框）
        for checkbox in self.channel_checkboxes.values():
            checkbox.blockSignals(True)  # 临时断开信号，避免触发更新
            checkbox.setChecked(False)
            checkbox.blockSignals(False)  # 恢复信号
        
        # 更新曲线显示
        self.update_curves()
    
    def on_well_selected(self, well_name):
        """孔位被选中时的处理"""
        # 如果well_name为空字符串，说明是清除选择
        if well_name == "":
            self.update_curves()
        else:
            self.update_curves()
        
        # 更新研判结果（根据选中的孔位）
        if self.selected_project:
            self.update_judgment_results()
    
    def on_channel_changed(self):
        """通道选择改变"""
        self.update_curves()
    
    def on_curve_type_changed(self):
        """曲线类型改变"""
        if self.amplification_radio.isChecked():
            self.curve_type = 'amplification'
        elif self.raw_radio.isChecked():
            self.curve_type = 'raw'
        self.update_curves()
    
    def on_project_changed(self, sender_name=None):
        """项目选择改变"""
        # 如果传入了发送者名称，说明是特定复选框触发的
        if sender_name:
            sender_cb = self.project_checkboxes.get(sender_name)
            if sender_cb and sender_cb.isChecked():
                # 如果当前复选框被选中，取消其他所有复选框的选中状态
                for name, cb in self.project_checkboxes.items():
                    if name != sender_name and cb.isChecked():
                        cb.blockSignals(True)
                        cb.setChecked(False)
                        cb.blockSignals(False)
                self.selected_project = sender_name
            else:
                # 如果当前复选框被取消选中，清空选择
                self.selected_project = None
        else:
            # 兼容旧逻辑：获取选中的项目（只允许选择一个）
            selected_projects = [name for name, cb in self.project_checkboxes.items() if cb.isChecked()]
            
            if len(selected_projects) > 1:
                # 如果选中了多个，只保留最后一个
                for name in selected_projects[:-1]:
                    self.project_checkboxes[name].blockSignals(True)
                    self.project_checkboxes[name].setChecked(False)
                    self.project_checkboxes[name].blockSignals(False)
                selected_projects = selected_projects[-1:]
            
            if selected_projects:
                self.selected_project = selected_projects[0]
            else:
                self.selected_project = None
        
        # 更新研判结果
        self.update_judgment_results()
    
    def clear_all_state(self):
        """清理所有状态（打开新文件时调用）"""
        # 清除孔板选择和数据
        if hasattr(self, 'plate_selector'):
            self.plate_selector.clear_selection()
            # 清除所有孔位的数据和显示
            if hasattr(self.plate_selector, 'well_data'):
                self.plate_selector.well_data.clear()
            # 重置所有按钮的文本和样式
            if hasattr(self.plate_selector, 'well_buttons'):
                default_style = self.plate_selector.get_default_button_style()
                for well_name, btn in self.plate_selector.well_buttons.items():
                    btn.setText("")
                    btn.setToolTip(f"孔位 {well_name}")
                    btn.setStyleSheet(default_style)
        
        # 清除项目选择
        if hasattr(self, 'project_checkboxes'):
            for checkbox in self.project_checkboxes.values():
                checkbox.blockSignals(True)
                checkbox.setChecked(False)
                checkbox.blockSignals(False)
        self.selected_project = None
        
        # 清空研判结果表格
        if hasattr(self, 'judgment_table'):
            self.judgment_table.setRowCount(0)
            self.judgment_table.setColumnCount(0)
        
        # 清空数据模型
        self.parsed_data = None
        self.data_model = None
        self.well_data_map = {}
        
        # 清空曲线显示
        if hasattr(self, 'figure'):
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, '请先打开Excel文件', 
                   ha='center', va='center', fontsize=14)
            ax.set_xticks([])
            ax.set_yticks([])
            if hasattr(self, 'canvas'):
                self.canvas.draw()
        
        # 清空信息显示
        if hasattr(self, 'curve_info'):
            self.curve_info.clear()
        if hasattr(self, 'info_text'):
            self.info_text.clear()
        if hasattr(self, 'data_table'):
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
    
    def open_file(self):
        """打开Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择Excel文件', '', 'Excel Files (*.xlsx *.xls)'
        )
        
        if file_path:
            # 清理之前的状态
            self.clear_all_state()
            
            self.current_file = file_path
            self.file_label.setText(f'文件: {os.path.basename(file_path)}')
            self.statusBar().showMessage('正在解析文件...')
            
            try:
                # 解析Excel文件
                parser = ExcelParser()
                self.parsed_data = parser.parse(file_path)
                
                # 转换为统一数据模型
                vendor_type = parser.detect_vendor(file_path)
                self.data_model = ConverterFactory.convert_data(self.parsed_data, vendor_type)
                
                # 调试信息
                print(f"=== 数据解析调试信息 ===")
                print(f"解析的数据键: {list(self.parsed_data.keys())}")
                if 'amplification_data' in self.parsed_data:
                    amp_data = self.parsed_data['amplification_data']
                    print(f"扩增数据形状: {amp_data.shape if not amp_data.empty else '空'}")
                    print(f"扩增数据列: {list(amp_data.columns) if not amp_data.empty else '无'}")
                    if not amp_data.empty:
                        print(f"扩增数据前5行:\n{amp_data.head()}")
                if 'well_data' in self.parsed_data:
                    print(f"孔位数据: {list(self.parsed_data['well_data'].keys())}")
                
                print(f"数据模型孔位数量: {len(self.data_model.wells)}")
                if self.data_model.wells:
                    print(f"孔位列表: {list(self.data_model.wells.keys())}")
                    first_well = list(self.data_model.wells.values())[0]
                    print(f"第一个孔位通道: {list(first_well.channels.keys())}")
                    if first_well.channels:
                        first_channel = list(first_well.channels.keys())[0]
                        print(f"第一个通道数据长度: {len(first_well.channels[first_channel])}")
                        print(f"第一个通道前5个值: {first_well.channels[first_channel][:5]}")
                else:
                    print("警告: 数据模型中没有孔位数据！")
                print(f"=== 调试信息结束 ===\n")
                
                # 显示数据
                self.display_data()
                self.update_plate_data()
                
                # 立即更新曲线显示
                self.update_curves()
                
                # 更新研判结果
                self.update_judgment_results()
                
                self.statusBar().showMessage('文件解析完成')
                self.export_btn.setEnabled(True)
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'解析文件失败:\n{str(e)}')
                self.statusBar().showMessage('解析失败')
    
    def display_data(self):
        """显示解析的数据"""
        if not self.parsed_data:
            return
        
        # 显示基本信息
        info = f"文件: {os.path.basename(self.current_file)}\n"
        info += f"工作表数量: {len(self.parsed_data.get('sheets', {}))}\n"
        
        # 显示实验信息
        if 'experiment_info' in self.parsed_data:
            info += "\n实验信息:\n"
            for key, value in self.parsed_data['experiment_info'].items():
                info += f"  {key}: {value}\n"
        
        self.info_text.setText(info)
        
        # 显示数据表格
        if 'amplification_data' in self.parsed_data:
            data = self.parsed_data['amplification_data']
            if not data.empty:
                self.display_table(data)
    
    def display_table(self, df):
        """在表格中显示数据"""
        self.data_table.setRowCount(len(df))
        self.data_table.setColumnCount(len(df.columns))
        self.data_table.setHorizontalHeaderLabels([str(col) for col in df.columns])
        
        for i in range(len(df)):
            for j in range(len(df.columns)):
                value = df.iloc[i, j]
                item = QTableWidgetItem(str(value) if pd.notna(value) else '')
                self.data_table.setItem(i, j, item)
    
    def update_plate_data(self):
        """更新孔板数据"""
        if not self.parsed_data:
            return
        
        self.well_data_map = {}
        
        # 优先使用解析器提取的孔位数据
        if 'well_data' in self.parsed_data:
            for well_name, well_info in self.parsed_data['well_data'].items():
                self.well_data_map[well_name] = {
                    'ct': well_info.get('ct'),
                    'data': well_info
                }
        
        # 如果amplification_data中有孔位信息，也提取
        if 'amplification_data' in self.parsed_data:
            data = self.parsed_data['amplification_data']
            if not data.empty:
                # 检查是否有孔位列
                if 'Well' in data.columns or '孔位' in data.columns:
                    well_col = 'Well' if 'Well' in data.columns else '孔位'
                    for idx, row in data.iterrows():
                        well_name = str(row[well_col]) if pd.notna(row[well_col]) else None
                        if well_name and well_name not in self.well_data_map:
                            # 提取Ct值等信息
                            ct_value = None
                            for col in ['Ct', 'CT', 'ct']:
                                if col in data.columns:
                                    ct_value = row[col] if pd.notna(row[col]) else None
                                    break
                            
                            self.well_data_map[well_name] = {
                                'ct': ct_value,
                                'data': row
                            }
        
        # 如果没有提取到孔位数据，生成示例数据用于演示
        if not self.well_data_map:
            rows = 8  # 96孔板：A-H
            cols = 12  # 96孔板：1-12
            
            import random
            for row_idx in range(rows):
                row_label = chr(65 + row_idx)
                for col in range(1, cols + 1):
                    well_name = f"{row_label}{col}"
                    # 随机生成一些Ct值用于演示
                    if random.random() > 0.3:  # 70%的孔有数据
                        ct = random.uniform(20, 40)
                        self.well_data_map[well_name] = {
                            'ct': ct,
                            'data': None
                        }
        
        # 更新孔板显示
        for well_name, data in self.well_data_map.items():
            if well_name in self.plate_selector.well_buttons:
                self.plate_selector.set_well_data(well_name, data)
    
    def update_curves(self):
        """更新曲线显示"""
        if not self.data_model:
            # 显示提示信息
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, '请先打开Excel文件', 
                   ha='center', va='center', fontsize=14)
            ax.set_xticks([])
            ax.set_yticks([])
            self.canvas.draw()
            return
        
        # 获取选中的通道
        selected_channels = [ch for ch, cb in self.channel_checkboxes.items() 
                           if cb.isChecked()]
        
        if not selected_channels:
            # 显示提示信息
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, '请至少选择一个通道', 
                   ha='center', va='center', fontsize=14)
            ax.set_xticks([])
            ax.set_yticks([])
            self.canvas.draw()
            return
        
        # 获取选中的孔位
        selected_wells = self.plate_selector.get_selected_wells()
        
        # 如果没有选中孔位，显示所有孔位
        if not selected_wells:
            selected_wells = list(self.data_model.wells.keys())
        
        # 如果没有孔位数据，显示提示
        if not selected_wells:
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, '没有可用的孔位数据', 
                   ha='center', va='center', fontsize=14)
            ax.set_xticks([])
            ax.set_yticks([])
            self.canvas.draw()
            return
        
        # 如果只显示一个孔位
        if not self.show_all_wells_check.isChecked() and selected_wells:
            selected_wells = [selected_wells[0]]
        
        # 绘制曲线
        try:
            visualizer = DataVisualizer()
            if self.curve_type == 'amplification':
                visualizer.plot_amplification_curves(
                    self.figure, self.data_model, selected_wells, selected_channels
                )
            else:
                visualizer.plot_raw_curves(
                    self.figure, self.data_model, selected_wells, selected_channels
                )
            
            self.canvas.draw()
            
            # 更新信息显示（已禁用）
            # self.update_curve_info(selected_wells, selected_channels)
        except Exception as e:
            # 显示错误信息
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, f'绘制曲线时出错:\n{str(e)}', 
                   ha='center', va='center', fontsize=12)
            ax.set_xticks([])
            ax.set_yticks([])
            self.canvas.draw()
            print(f"绘制曲线错误: {e}")  # 调试信息
    
    def update_curve_info(self, well_names: List[str], channel_names: List[str]):
        """更新曲线信息显示"""
        if not self.data_model:
            return
        
        info_text = f"显示信息:\n"
        info_text += f"曲线类型: {'扩增曲线' if self.curve_type == 'amplification' else '原始曲线'}\n"
        info_text += f"孔位: {', '.join(well_names) if well_names else '全部'}\n"
        info_text += f"通道: {', '.join(channel_names)}\n\n"
        
        # 显示Ct值信息
        for well_name in well_names:
            well = self.data_model.get_well(well_name)
            if well and well.ct_values:
                info_text += f"{well_name} Ct值:\n"
                for channel, ct in well.ct_values.items():
                    info_text += f"  {channel}: {ct:.2f}\n"
        
        self.curve_info.setText(info_text)
    
    def export_results(self):
        """导出分析结果"""
        if not self.parsed_data:
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, '保存结果', '', 'Excel Files (*.xlsx)'
        )
        
        if file_path:
            try:
                # 导出逻辑
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    if 'amplification_data' in self.parsed_data:
                        self.parsed_data['amplification_data'].to_excel(
                            writer, sheet_name='分析结果', index=False
                        )
                QMessageBox.information(self, '成功', '结果已导出')
            except Exception as e:
                QMessageBox.critical(self, '错误', f'导出失败:\n{str(e)}')
    
    def update_judgment_results(self):
        """更新研判结果显示"""
        if not self.data_model or not self.selected_project:
            # 清空表格
            self.judgment_table.setRowCount(0)
            return
        
        # 获取项目配置
        if self.selected_project not in projects_data:
            self.judgment_table.setRowCount(0)
            return
        
        project_config = projects_data[self.selected_project]
        
        # 筛选出有target的通道
        valid_channels = []
        for ch_name in project_channel_names:
            if ch_name in project_config:
                ch_config = project_config[ch_name]
                target = ch_config.get('target', '')
                # 只包含有target且target不为空的通道
                if target and target.strip() and target.strip() != '\\' and target.strip() != '/':
                    valid_channels.append(ch_name)
        
        # 如果没有有效通道，清空表格
        if not valid_channels:
            self.judgment_table.setRowCount(0)
            self.judgment_table.setColumnCount(0)
            return
        
        # 设置表格列数和列标题
        column_count = 2 + len(valid_channels) + 1  # Well + 项目名 + 通道列 + 判读结果
        self.judgment_table.setColumnCount(column_count)
        headers = ['Well', '项目名'] + valid_channels + ['判读结果']
        self.judgment_table.setHorizontalHeaderLabels(headers)
        
        # 获取在孔板中选择的孔位
        selected_wells = self.plate_selector.get_selected_wells()
        
        # 如果没有选中任何孔位，显示所有孔位
        if not selected_wells:
            all_wells = list(self.data_model.wells.keys())
        else:
            # 只显示选中的孔位
            all_wells = selected_wells
        
        if not all_wells:
            self.judgment_table.setRowCount(0)
            return
        
        # 准备结果数据
        results = []
        
        for well_name in sorted(all_wells):
            well = self.data_model.get_well(well_name)
            if not well:
                continue
            
            # 获取各通道的CT值（只处理有效通道）
            ct_values = {}
            positive_targets = []
            
            for ch_name in valid_channels:
                # 获取CT值
                ct_value = well.ct_values.get(ch_name, None)
                ct_values[ch_name] = ct_value
                
                # 判断是否阳性
                if ch_name in project_config:
                    ch_config = project_config[ch_name]
                    threshold = ch_config.get('threshold', None)
                    target = ch_config.get('target', '')
                    
                    if ct_value is not None and threshold is not None:
                        # CT值小于阈值则为阳性（CT值越小，扩增越早，越可能是阳性）
                        if ct_value < threshold:
                            positive_targets.append(target if target else ch_name)
            
            # 所有孔位都添加到结果中，即使没有CT值
            results.append({
                'well': well_name,
                'project': self.selected_project,
                'ct_values': ct_values,
                'positive_targets': positive_targets
            })
        
        # 更新表格
        self.judgment_table.setRowCount(len(results))
        
        for row_idx, result in enumerate(results):
            # Well
            self.judgment_table.setItem(row_idx, 0, QTableWidgetItem(result['well']))
            
            # 项目名
            self.judgment_table.setItem(row_idx, 1, QTableWidgetItem(result['project']))
            
            # 各通道CT值（只显示有效通道）
            for col_idx, ch_name in enumerate(valid_channels, 2):
                ct_value = result['ct_values'].get(ch_name)
                if ct_value is not None:
                    item = QTableWidgetItem(f"{ct_value:.2f}")
                    # 根据是否阳性设置颜色
                    if ch_name in project_config:
                        ch_config = project_config[ch_name]
                        threshold = ch_config.get('threshold', None)
                        if threshold is not None and ct_value < threshold:
                            item.setBackground(QColor(255, 200, 200))  # 浅红色 - 阳性
                        else:
                            item.setBackground(QColor(200, 255, 200))  # 浅绿色 - 阴性
                    self.judgment_table.setItem(row_idx, col_idx, item)
                else:
                    self.judgment_table.setItem(row_idx, col_idx, QTableWidgetItem("N/A"))
            
            # 判读结果（阳性targets）
            result_col_idx = 2 + len(valid_channels)  # 判读结果列索引
            if result['positive_targets']:
                result_text = ", ".join(result['positive_targets'])
                item = QTableWidgetItem(result_text)
                item.setBackground(QColor(255, 200, 200))  # 浅红色
                self.judgment_table.setItem(row_idx, result_col_idx, item)
            else:
                self.judgment_table.setItem(row_idx, result_col_idx, QTableWidgetItem("阴性"))
        
        # 调整列宽
        self.judgment_table.resizeColumnsToContents()


def main():
    app = QApplication(sys.argv)
    window = PCRAnalyzerApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
