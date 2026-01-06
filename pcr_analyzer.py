"""
PCR结果分析软件
支持解析不同厂商的Excel格式，展示PCR扩增结果
"""
import sys
import os
import shutil
from pathlib import Path
from typing import List
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QFileDialog, QTableWidget, 
                             QTableWidgetItem, QLabel, QMessageBox, QTabWidget,
                             QTextEdit, QSplitter, QGroupBox, QComboBox, QCheckBox, QRadioButton, QButtonGroup, QListWidget, QLineEdit, QStyle)
from PyQt5.QtGui import QFontMetrics
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

# 通道名称列表
project_channel_names = ['FAM', 'VIC', 'CY5', 'ROX']


def load_projects_from_excel(file_path):
    """
    从Excel文件加载项目数据
    支持两种格式：
    1. 每个项目一行，列格式：项目名称, project_id, FAM_target, FAM_threshold, VIC_target, VIC_threshold, ...
    2. 每个项目多行，每行一个通道
    """
    try:
        # 根据文件扩展名选择引擎
        file_ext = Path(file_path).suffix.lower()
        if file_ext == '.xls':
            # .xls 文件使用 xlrd 引擎
            try:
                df = pd.read_excel(file_path, sheet_name=0, header=None, engine='xlrd')
            except ImportError:
                raise ImportError("读取 .xls 文件需要安装 xlrd>=2.0.1，请运行: pip install xlrd>=2.0.1")
        else:
            # .xlsx 文件使用 openpyxl 引擎
            df = pd.read_excel(file_path, sheet_name=0, header=None, engine='openpyxl')
        
        projects = {}
        channels = ['FAM', 'VIC', 'CY5', 'ROX']
        
        # 查找表头行（可能有多行表头）
        header_row = None
        for idx, row in df.iterrows():
            row_str = ' '.join([str(x).upper() for x in row if pd.notna(x)])
            if '项目' in row_str or 'PROJECT' in row_str or 'FAM' in row_str:
                header_row = idx
                break
        
        if header_row is None:
            # 如果没有找到表头，假设第一行是表头
            header_row = 0
        
        # 读取表头，确定列索引
        # 先读取主表头行
        header = df.iloc[header_row].values
        
        # 查找各列的索引
        col_map = {}
        for i, val in enumerate(header):
            if pd.notna(val):
                val_str = str(val).upper()
                # 先检查是否是项目名称列
                if '项目' in val_str or 'PROJECT' in val_str or 'NAME' in val_str:
                    # 如果已经包含"编号"或"ID"，可能是项目编号列而不是名称列
                    if '编号' not in val_str and 'ID' not in val_str:
                        col_map['project_name'] = i
                # 检查是否是项目编号列（更宽松的条件）
                elif ('编号' in val_str or 'ID' in val_str) and 'project_id' not in col_map:
                    # 排除通道相关的列（FAM, VIC, CY5, ROX等）
                    is_channel_col = False
                    for ch in channels:
                        if ch in val_str:
                            is_channel_col = True
                            break
                    # 排除target和threshold相关的列
                    if not is_channel_col and 'TARGET' not in val_str and 'THRESHOLD' not in val_str and '目标' not in val_str and '阈值' not in val_str and '靶标' not in val_str:
                        col_map['project_id'] = i
                else:
                    # 检查通道相关列（更灵活的匹配）
                    for ch in channels:
                        if ch in val_str:
                            # 检查是否是target列
                            if 'TARGET' in val_str or '目标' in val_str or '靶标' in val_str:
                                col_map[f'{ch}_target'] = i
                            # 检查是否是threshold列
                            elif 'THRESHOLD' in val_str or '阈值' in val_str or 'TH' in val_str:
                                col_map[f'{ch}_threshold'] = i
        
        # 如果主表头行没有找到项目编号列，检查下一行（可能是多行表头）
        if 'project_id' not in col_map and header_row + 1 < len(df):
            next_header = df.iloc[header_row + 1].values
            for i, val in enumerate(next_header):
                if pd.notna(val):
                    val_str = str(val).upper()
                    # 检查是否是项目编号列
                    if ('编号' in val_str or 'ID' in val_str or '产品' in val_str) and 'project_id' not in col_map:
                        # 排除通道相关的列
                        is_channel_col = False
                        for ch in channels:
                            if ch in val_str:
                                is_channel_col = True
                                break
                        # 排除target和threshold相关的列
                        if not is_channel_col and 'TARGET' not in val_str and 'THRESHOLD' not in val_str and '目标' not in val_str and '阈值' not in val_str and '靶标' not in val_str:
                            col_map['project_id'] = i
                            break
            
            # 也检查下一行中的通道相关列（可能是多行表头，通道名在第一行，target/threshold在第二行）
            for i, val in enumerate(next_header):
                if pd.notna(val):
                    val_str = str(val).upper()
                    # 检查是否是target或threshold列（不包含通道名，可能是第二行表头）
                    if 'TARGET' in val_str or '目标' in val_str or '靶标' in val_str:
                        # 查找对应的通道（通过列位置推断，或者检查上一行对应位置）
                        if i < len(header):
                            prev_val = header[i] if pd.notna(header[i]) else ''
                            prev_val_str = str(prev_val).upper()
                            for ch in channels:
                                if ch in prev_val_str:
                                    if f'{ch}_target' not in col_map:
                                        col_map[f'{ch}_target'] = i
                                    break
                    elif 'THRESHOLD' in val_str or '阈值' in val_str or 'TH' in val_str:
                        # 查找对应的通道
                        if i < len(header):
                            prev_val = header[i] if pd.notna(header[i]) else ''
                            prev_val_str = str(prev_val).upper()
                            for ch in channels:
                                if ch in prev_val_str:
                                    if f'{ch}_threshold' not in col_map:
                                        col_map[f'{ch}_threshold'] = i
                                    break
        
        # 如果找不到通道列映射，尝试按位置推断
        # 检查是否所有通道的target和threshold都没有找到
        missing_channels = []
        for ch in channels:
            if f'{ch}_target' not in col_map or f'{ch}_threshold' not in col_map:
                missing_channels.append(ch)
        
        if missing_channels:
            # 根据实际数据格式推断：
            # 列2-5: FAM, VIC, CY5, ROX 的target（目标）
            # 列6-9: FAM, VIC, CY5, ROX 的threshold（阈值）
            if 'project_name' in col_map and 'project_id' in col_map:
                # 从列2开始是target，列6开始是threshold
                target_start_col = 2
                threshold_start_col = 6
                for i, ch in enumerate(missing_channels):
                    if f'{ch}_target' not in col_map:
                        col_map[f'{ch}_target'] = target_start_col + i
                    if f'{ch}_threshold' not in col_map:
                        col_map[f'{ch}_threshold'] = threshold_start_col + i
        
        # 确定数据起始行：如果项目编号列在下一行找到，数据从header_row+2开始，否则从header_row+1开始
        data_start_row = header_row + 1
        if 'project_id' in col_map:
            # 检查项目编号列是否在下一行（多行表头的情况）
            if header_row + 1 < len(df):
                next_header = df.iloc[header_row + 1].values
                if col_map['project_id'] < len(next_header):
                    next_val = next_header[col_map['project_id']]
                    if pd.notna(next_val):
                        val_str = str(next_val).upper()
                        # 如果下一行对应位置有值且包含"编号"或"产品"，说明是表头的一部分
                        if '编号' in val_str or 'ID' in val_str or '产品' in val_str:
                            data_start_row = header_row + 2
        
        # 读取数据行
        for idx in range(data_start_row, len(df)):
            row = df.iloc[idx]
            
            # 获取项目名称
            if 'project_name' not in col_map:
                continue
            project_name_col = col_map['project_name']
            if project_name_col >= len(row) or pd.isna(row.iloc[project_name_col]):
                continue
            
            project_name = str(row.iloc[project_name_col]).strip()
            if not project_name or project_name == 'nan':
                continue
            
            # 获取project_id
            project_id = ''
            if 'project_id' in col_map and col_map['project_id'] < len(row):
                project_id_val = row.iloc[col_map['project_id']]
                if pd.notna(project_id_val):
                    project_id = str(project_id_val).strip()
            
            # 初始化项目数据
            if project_name not in projects:
                projects[project_name] = {
                    'project_id': project_id,
                }
            
            # 读取各通道数据
            for ch in channels:
                if ch not in projects[project_name]:
                    projects[project_name][ch] = {}
                
                # 读取target
                target_key = f'{ch}_target'
                if target_key in col_map and col_map[target_key] < len(row):
                    target_val = row.iloc[col_map[target_key]]
                    if pd.notna(target_val):
                        target_str = str(target_val).strip()
                        if target_str and target_str != 'nan':
                            projects[project_name][ch]['target'] = target_str
                
                # 读取threshold
                threshold_key = f'{ch}_threshold'
                if threshold_key in col_map and col_map[threshold_key] < len(row):
                    threshold_val = row.iloc[col_map[threshold_key]]
                    if pd.notna(threshold_val):
                        try:
                            threshold = float(threshold_val)
                            projects[project_name][ch]['threshold'] = threshold
                        except:
                            pass
        
        return projects
    
    except Exception as e:
        print(f"读取项目Excel文件失败: {e}")
        return {}


def get_base_directory():
    """获取程序所在目录（支持打包成exe）"""
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe，使用exe所在目录
        return Path(sys.executable).parent
    else:
        # 如果是脚本运行，使用脚本所在目录
        return Path(__file__).parent


def load_projects_data():
    """加载项目数据，从目录下的projects.xls或projects.xlsx读取，如果不存在则返回空"""
    default_channels = ['FAM', 'VIC', 'CY5', 'ROX']
    
    # 尝试从目录下的文件读取
    base_dir = get_base_directory()
    project_files = [
        base_dir / 'projects.xlsx',
        base_dir / 'projects.xls',
    ]
    
    for file_path in project_files:
        if file_path.exists():
            try:
                loaded_projects = load_projects_from_excel(str(file_path))
                if loaded_projects:
                    print(f"从 {file_path.name} 加载了 {len(loaded_projects)} 个项目")
                    return loaded_projects, default_channels
            except Exception as e:
                print(f"读取 {file_path.name} 失败: {e}")
    
    # 如果没有找到文件或读取失败，返回空字典（不显示项目）
    return {}, default_channels


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
        self.selected_projects = []  # 选中的项目名称列表（支持多选）
        self.judgment_results = []  # 结果判读列表
        
        # 加载项目数据
        self.projects_data, self.project_channel_names = load_projects_data()
        
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
        # 设置孔板选择器的尺寸策略，保持矩形形状
        # 96孔板是8行12列，计算实际需要的宽度：
        # 行标签25 + 12列按钮(12*40) + 间距(11*2) + 边距 ≈ 550像素
        # 高度：标题 + 列标签 + 8行按钮(8*40) + 间距 + 边距 ≈ 380像素
        self.plate_selector.setMinimumSize(550, 380)
        self.plate_selector.setMaximumSize(600, 420)
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
        
        # 全选按钮
        select_all_btn = QPushButton('全选')
        select_all_btn.clicked.connect(self.select_all_wells_and_channels)
        left_layout.addWidget(select_all_btn)
        
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
        
        # 曲线类型选择（只保留单选按钮，不显示组框和标签）
        self.curve_type_group = QButtonGroup()
        self.amplification_radio = QRadioButton('扩增曲线')
        self.amplification_radio.setChecked(True)
        self.raw_radio = QRadioButton('原始曲线')
        self.curve_type_group.addButton(self.amplification_radio, 0)
        self.curve_type_group.addButton(self.raw_radio, 1)
        self.amplification_radio.toggled.connect(self.on_curve_type_changed)
        self.raw_radio.toggled.connect(self.on_curve_type_changed)
        control_layout.addWidget(self.amplification_radio)
        control_layout.addWidget(self.raw_radio)
        
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
        
        # 保存项目组布局的引用，以便刷新项目列表
        self.project_group_layout = project_group_layout
        
        # 导入项目按钮
        import_project_btn = QPushButton('导入项目')
        import_project_btn.clicked.connect(self.import_projects)
        project_group_layout.addWidget(import_project_btn)
        
        # 搜索框
        search_label = QLabel("搜索:")
        project_group_layout.addWidget(search_label)
        self.project_search_box = QLineEdit()
        self.project_search_box.setPlaceholderText("输入项目名称或产品编号...")
        self.project_search_box.textChanged.connect(self.on_project_search_changed)
        project_group_layout.addWidget(self.project_search_box)
        
        # 项目复选框列表
        self.project_checkboxes = {}
        self.all_project_checkboxes = {}  # 保存所有项目复选框的引用
        self.current_page = 1  # 当前页码
        self.items_per_page = 20  # 每页显示的项目数
        
        # 分页控件（放在项目列表下方，先创建以便refresh_project_list可以使用）
        pagination_layout = QHBoxLayout()
        pagination_layout.addStretch()
        
        self.prev_page_btn = QPushButton()
        self.prev_page_btn.setIcon(self.style().standardIcon(QStyle.SP_ArrowLeft))
        self.prev_page_btn.setFixedSize(30, 30)
        self.prev_page_btn.clicked.connect(self.go_to_prev_page)
        self.prev_page_btn.setEnabled(False)
        pagination_layout.addWidget(self.prev_page_btn)
        
        self.page_label = QLabel('1/1')
        pagination_layout.addWidget(self.page_label)
        
        self.next_page_btn = QPushButton()
        self.next_page_btn.setIcon(self.style().standardIcon(QStyle.SP_ArrowRight))
        self.next_page_btn.setFixedSize(30, 30)
        self.next_page_btn.clicked.connect(self.go_to_next_page)
        self.next_page_btn.setEnabled(False)
        pagination_layout.addWidget(self.next_page_btn)
        
        pagination_layout.addStretch()
        project_group_layout.addLayout(pagination_layout)
        
        # 在分页控件创建后再刷新项目列表
        self.refresh_project_list()
        
        project_group_layout.addStretch()
        project_layout.addWidget(project_group)
        project_layout.addStretch()
        
        # 将三个面板添加到水平分割器
        top_splitter.addWidget(left_panel)
        top_splitter.addWidget(right_panel)
        top_splitter.addWidget(project_panel)
        # 设置各面板的拉伸因子和最小宽度
        top_splitter.setStretchFactor(0, 1)  # 左侧孔板
        top_splitter.setStretchFactor(1, 3)  # 中间曲线（增加宽度）
        top_splitter.setStretchFactor(2, 1)  # 右侧项目
        # 设置最小宽度，确保项目面板有足够空间（增加以容纳滚动条）
        project_panel.setMinimumWidth(350)
        # 增加左侧面板最小宽度，确保孔板完整显示（96孔板需要约550像素宽度）
        left_panel.setMinimumWidth(600)
        
        # 添加到主布局（第一行）
        main_layout.addWidget(top_splitter, 2)  # 占据2倍空间
        
        # 第二行：结果判读显示区域（单独一行，占据整个宽度）
        judgment_group = QGroupBox()
        judgment_layout = QVBoxLayout(judgment_group)
        
        # 创建标题栏（标题和导出按钮在同一行，紧挨着）
        title_layout = QHBoxLayout()
        title_label = QLabel("结果判读")
        title_label.setStyleSheet("font-weight: bold;")
        title_layout.addWidget(title_label)
        # 添加一些间距
        title_layout.addSpacing(10)
        self.export_judgment_btn = QPushButton('导出结果')
        self.export_judgment_btn.clicked.connect(self.export_judgment_results)
        title_layout.addWidget(self.export_judgment_btn)
        title_layout.addStretch()  # 让标题和按钮靠左，其余空间在右边
        judgment_layout.addLayout(title_layout)
        
        # 结果显示表格（列数和列标题会根据选中的项目动态设置）
        self.judgment_table = QTableWidget()
        # 初始列数会在update_judgment_results中动态设置
        # 设置最小高度，让结果表格有足够的显示空间
        self.judgment_table.setMinimumHeight(250)
        judgment_layout.addWidget(self.judgment_table)
        
        # 添加到主布局（第二行）
        main_layout.addWidget(judgment_group, 1)  # 占据1倍空间
        
        return widget
    
    def select_all_wells_and_channels(self):
        """全选所有孔位和通道"""
        # 全选所有通道
        for checkbox in self.channel_checkboxes.values():
            checkbox.blockSignals(True)
            checkbox.setChecked(True)
            checkbox.blockSignals(False)
        
        # 更新选中的通道列表
        self.selected_channels = list(self.channel_checkboxes.keys())
        
        # 全选所有孔位（通过plate_selector的toggle_select_all方法）
        if hasattr(self, 'plate_selector') and hasattr(self.plate_selector, 'toggle_select_all'):
            # 检查是否已全部选中
            all_selected = len(self.plate_selector.selected_wells) == len(self.plate_selector.well_buttons)
            if not all_selected:
                self.plate_selector.toggle_select_all()
        
        # 更新曲线显示
        self.update_curves()
    
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
        
        # 更新结果判读（根据选中的孔位）
        if self.selected_projects:
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
        """项目选择改变（支持多选）"""
        # 获取所有选中的项目
        self.selected_projects = [name for name, cb in self.project_checkboxes.items() if cb.isChecked()]
        
        # 更新结果判读
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
        self.selected_projects = []
        
        # 清空结果判读表格
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
                
                # 更新孔板数据
                self.update_plate_data()
                
                # 立即更新曲线显示
                self.update_curves()
                
                # 更新结果判读
                self.update_judgment_results()
                
                self.statusBar().showMessage('文件解析完成')
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'解析文件失败:\n{str(e)}')
                self.statusBar().showMessage('解析失败')
    
    def update_plate_data(self):
        """更新孔板数据"""
        if not self.data_model:
            return
        
        self.well_data_map = {}
        
        # 从data_model中获取CT值
        for well_name, well in self.data_model.wells.items():
            # 获取所有通道的CT值，选择最小的CT值显示（如果有多个通道）
            ct_value = None
            if well.ct_values:
                # 获取所有有效的CT值
                valid_cts = [ct for ct in well.ct_values.values() if ct is not None and pd.notna(ct)]
                if valid_cts:
                    # 使用最小的CT值（通常表示最早检测到信号）
                    ct_value = min(valid_cts)
            
            self.well_data_map[well_name] = {
                'ct': ct_value,
                'data': well
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
        
        # 始终显示所有选中的孔位（不再限制为只显示一个）
        
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
    
    def on_project_search_changed(self, text):
        """项目搜索框内容改变时的处理"""
        self.current_page = 1  # 搜索时重置到第一页
        self.filter_project_list(text)
    
    def get_filtered_projects(self, search_text=""):
        """获取过滤后的项目列表"""
        if not self.projects_data:
            return []
        
        search_text = search_text.strip().upper()
        filtered_projects = []
        
        for project_name in sorted(self.projects_data.keys()):
            if not search_text:
                # 如果没有搜索文本，包含所有项目
                filtered_projects.append(project_name)
            else:
                # 检查项目名称是否匹配
                project_name_match = search_text in project_name.upper()
                
                # 检查产品编号是否匹配
                project_id_match = False
                if project_name in self.projects_data:
                    project_id = self.projects_data[project_name].get('project_id', '')
                    if project_id and search_text in str(project_id).upper():
                        project_id_match = True
                
                # 如果项目名称或产品编号匹配，包含该项目
                if project_name_match or project_id_match:
                    filtered_projects.append(project_name)
        
        return filtered_projects
    
    def filter_project_list(self, search_text=""):
        """根据搜索文本和分页过滤项目列表"""
        if not hasattr(self, 'all_project_checkboxes'):
            return
        
        # 获取过滤后的项目列表
        filtered_projects = self.get_filtered_projects(search_text)
        
        # 计算总页数
        total_pages = max(1, (len(filtered_projects) + self.items_per_page - 1) // self.items_per_page)
        
        # 确保当前页在有效范围内
        if self.current_page > total_pages:
            self.current_page = total_pages
        if self.current_page < 1:
            self.current_page = 1
        
        # 计算当前页的项目范围
        start_idx = (self.current_page - 1) * self.items_per_page
        end_idx = start_idx + self.items_per_page
        current_page_projects = filtered_projects[start_idx:end_idx]
        
        # 显示或隐藏项目复选框
        for project_name, checkbox in self.all_project_checkboxes.items():
            checkbox.setVisible(project_name in current_page_projects)
        
        # 更新分页控件
        self.update_pagination_controls(total_pages, len(filtered_projects))
    
    def update_pagination_controls(self, total_pages, total_items):
        """更新分页控件状态"""
        if not hasattr(self, 'page_label') or not self.page_label:
            return  # 分页控件还未创建
        self.page_label.setText(f'{self.current_page}/{total_pages}')
        self.prev_page_btn.setEnabled(self.current_page > 1)
        self.next_page_btn.setEnabled(self.current_page < total_pages)
    
    def go_to_prev_page(self):
        """跳转到上一页"""
        if self.current_page > 1:
            self.current_page -= 1
            if hasattr(self, 'project_search_box'):
                self.filter_project_list(self.project_search_box.text())
    
    def go_to_next_page(self):
        """跳转到下一页"""
        if hasattr(self, 'project_search_box'):
            search_text = self.project_search_box.text()
            filtered_projects = self.get_filtered_projects(search_text)
            total_pages = max(1, (len(filtered_projects) + self.items_per_page - 1) // self.items_per_page)
            if self.current_page < total_pages:
                self.current_page += 1
                self.filter_project_list(search_text)
    
    def refresh_project_list(self):
        """刷新项目列表显示"""
        # 清除现有的复选框和标签（保留导入按钮和搜索框）
        widgets_to_remove = []
        for i in range(self.project_group_layout.count()):
            item = self.project_group_layout.itemAt(i)
            if item:
                widget = item.widget()
                # 保留导入按钮（索引0）、搜索标签、搜索框
                if widget and isinstance(widget, QCheckBox):
                    widgets_to_remove.append(widget)
                elif widget and isinstance(widget, QLabel) and widget.text() == "未找到项目数据":
                    widgets_to_remove.append(widget)
        
        for widget in widgets_to_remove:
            widget.setParent(None)
            self.project_group_layout.removeWidget(widget)
        
        self.project_checkboxes.clear()
        self.all_project_checkboxes.clear()
        
        # 添加新的项目复选框（跳过导入按钮、搜索标签和搜索框）
        if self.projects_data:
            for project_name in sorted(self.projects_data.keys()):
                checkbox = QCheckBox(project_name)
                # 创建包装函数来传递项目名称，避免lambda闭包问题
                def make_handler(name):
                    def handler(checked):
                        self.on_project_changed(name)
                    return handler
                checkbox.stateChanged.connect(make_handler(project_name))
                self.project_checkboxes[project_name] = checkbox
                self.all_project_checkboxes[project_name] = checkbox
                # 插入到搜索框之后（索引3：导入按钮0，搜索标签1，搜索框2）
                self.project_group_layout.insertWidget(3, checkbox)
            
            # 应用当前的搜索过滤和分页
            if hasattr(self, 'project_search_box'):
                self.current_page = 1  # 刷新时重置到第一页
                self.filter_project_list(self.project_search_box.text())
        else:
            no_project_label = QLabel("未找到项目数据")
            self.project_group_layout.insertWidget(3, no_project_label)
    
    def import_projects(self):
        """导入项目Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择项目Excel文件', '', 'Excel Files (*.xlsx *.xls)'
        )
        
        if file_path:
            try:
                loaded_projects = load_projects_from_excel(file_path)
                if loaded_projects:
                    # 直接使用Excel数据
                    self.projects_data = loaded_projects
                    # 刷新项目列表
                    self.refresh_project_list()
                    # 清空当前选择
                    self.selected_projects = []
                    # 更新结果判读
                    self.update_judgment_results()
                    
                    # 将导入的文件保存到当前文件夹下，作为默认项目模板
                    try:
                        base_dir = get_base_directory()
                        target_file = base_dir / 'projects.xls'
                        
                        # 如果源文件是.xlsx，保存为.xlsx；如果是.xls，保存为.xls
                        source_ext = Path(file_path).suffix.lower()
                        if source_ext == '.xlsx':
                            target_file = base_dir / 'projects.xlsx'
                        
                        # 复制文件
                        shutil.copy2(file_path, target_file)
                        QMessageBox.information(self, '成功', 
                            f'成功导入 {len(loaded_projects)} 个项目\n'
                            f'已保存为默认项目模板: {target_file.name}')
                    except Exception as save_error:
                        # 如果保存失败，仍然显示导入成功，但提示保存失败
                        QMessageBox.warning(self, '导入成功但保存失败', 
                            f'成功导入 {len(loaded_projects)} 个项目\n'
                            f'但保存为默认模板失败: {str(save_error)}')
                else:
                    QMessageBox.warning(self, '警告', '未能从文件中读取到项目数据')
            except Exception as e:
                QMessageBox.critical(self, '错误', f'导入项目失败:\n{str(e)}')
    
    def update_judgment_results(self):
        """更新结果判读显示（支持多项目）"""
        if not self.data_model:
            # 如果没有数据模型，显示提示
            self.judgment_table.setRowCount(0)
            self.judgment_table.setColumnCount(1)
            self.judgment_table.setHorizontalHeaderLabels(['提示'])
            self.judgment_table.setRowCount(1)
            self.judgment_table.setItem(0, 0, QTableWidgetItem("请先打开Excel文件"))
            return
        
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
        
        # 如果没有选择项目，不显示判读结果
        if not self.selected_projects:
            self.judgment_table.setRowCount(0)
            self.judgment_table.setColumnCount(0)
            return
        
        # 收集所有项目的有效通道（合并所有项目的通道）
        all_valid_channels = set()
        projects_configs = {}
        
        # 如果有选择项目，使用项目配置来确定有效通道
        for project_name in self.selected_projects:
            if project_name not in self.projects_data:
                # 如果项目不在数据中，跳过该项目
                print(f"警告: 项目 '{project_name}' 不在项目数据中")
                continue
            
            project_config = self.projects_data[project_name]
            projects_configs[project_name] = project_config
            
            # 筛选出有配置的通道（只要通道存在于项目配置中，就显示该通道的CT值）
            for ch_name in self.project_channel_names:
                if ch_name in project_config:
                    # 只要通道存在于项目配置中，就添加该通道（即使没有target或threshold）
                    all_valid_channels.add(ch_name)
        
        # 如果没有有效通道，清空表格并显示提示
        if not all_valid_channels:
            self.judgment_table.setRowCount(0)
            self.judgment_table.setColumnCount(1)
            self.judgment_table.setHorizontalHeaderLabels(['提示'])
            self.judgment_table.setRowCount(1)
            if not projects_configs:
                self.judgment_table.setItem(0, 0, QTableWidgetItem("所选项目不在项目数据中，请先导入项目"))
            else:
                self.judgment_table.setItem(0, 0, QTableWidgetItem("所选项目没有有效的通道配置"))
            return
        
        # 按自定义顺序排序通道：FAM、VIC、ROX、CY5
        channel_order = ['FAM', 'VIC', 'ROX', 'CY5']
        # 先按自定义顺序排序，然后添加其他不在列表中的通道（按字母顺序）
        valid_channels = []
        for ch in channel_order:
            if ch in all_valid_channels:
                valid_channels.append(ch)
        # 添加其他通道（按字母顺序）
        remaining_channels = sorted([ch for ch in all_valid_channels if ch not in channel_order])
        valid_channels.extend(remaining_channels)
        
        # 设置表格列数和列标题
        column_count = 4 + len(valid_channels) + 1  # Well + 样本 + 项目名 + 产品编号 + 通道列 + 判读结果
        self.judgment_table.setColumnCount(column_count)
        # 给每个通道名称添加"(CT)"后缀
        channel_headers = [f"{ch}(CT)" for ch in valid_channels]
        headers = ['Well', '样本', '项目名', '产品编号'] + channel_headers + ['判读结果']
        self.judgment_table.setHorizontalHeaderLabels(headers)
        
        # 准备结果数据（每个孔位 × 每个项目）
        results = []
        
        def get_ct_value(well, channel_name):
            """获取通道的CT值，VIC和HEX等价，如果VIC没有数据则取HEX的数据"""
            ct_value = well.ct_values.get(channel_name, None)
            # 如果VIC没有数据，尝试从HEX获取
            if channel_name == 'VIC' and ct_value is None:
                ct_value = well.ct_values.get('HEX', None)
            return ct_value
        
        for well_name in sorted(all_wells):
            well = self.data_model.get_well(well_name)
            if not well:
                continue
            
            # 为每个选中的项目生成结果
            for project_name in self.selected_projects:
                if project_name not in projects_configs:
                    continue
                
                project_config = projects_configs[project_name]
                
                # 获取各通道的CT值（只处理有效通道）
                ct_values = {}
                positive_targets = []
                
                for ch_name in valid_channels:
                    # 获取CT值（VIC和HEX等价）
                    ct_value = get_ct_value(well, ch_name)
                    ct_values[ch_name] = ct_value
                    
                    # 判断是否阳性（只判断当前项目配置中包含的通道）
                    if ch_name in project_config:
                        ch_config = project_config[ch_name]
                        threshold = ch_config.get('threshold', None)
                        target = ch_config.get('target', '')
                        
                        if ct_value is not None and threshold is not None:
                            # CT值小于阈值则为阳性（CT值越小，扩增越早，越可能是阳性）
                            if ct_value < threshold:
                                # 只添加有效的target，过滤掉空值、"\"、"/"等无效target
                                if target and target.strip() and target.strip() != '\\' and target.strip() != '/':
                                    positive_targets.append(target)
                                # 如果target无效，不添加到列表中（不显示通道名）
                
                # 获取产品编号
                project_id = project_config.get('project_id', '')
                
                # 获取样本名称
                sample_name = well.metadata.get('sample_name', '') if well.metadata else ''
                
                # 添加到结果中
                results.append({
                    'well': well_name,
                    'sample_name': sample_name,
                    'project': project_name,
                    'project_id': project_id,
                    'ct_values': ct_values,
                    'positive_targets': positive_targets,
                    'project_config': project_config
                })
        
        # 更新表格
        self.judgment_table.setRowCount(len(results))
        
        for row_idx, result in enumerate(results):
            # Well
            self.judgment_table.setItem(row_idx, 0, QTableWidgetItem(result['well']))
            
            # 样本名称
            sample_name = result.get('sample_name', '')
            self.judgment_table.setItem(row_idx, 1, QTableWidgetItem(str(sample_name) if sample_name else ''))
            
            # 项目名
            self.judgment_table.setItem(row_idx, 2, QTableWidgetItem(result['project']))
            
            # 产品编号
            project_id = result.get('project_id', '')
            self.judgment_table.setItem(row_idx, 3, QTableWidgetItem(str(project_id) if project_id else ''))
            
            # 各通道CT值（只显示有效通道）
            project_config = result['project_config']
            for col_idx, ch_name in enumerate(valid_channels, 4):
                ct_value = result['ct_values'].get(ch_name)
                # 如果VIC没有数据，尝试从HEX获取（在显示时也需要处理）
                if ct_value is None and ch_name == 'VIC':
                    well = self.data_model.get_well(result['well'])
                    if well:
                        ct_value = well.ct_values.get('HEX', None)
                        # 更新结果中的CT值，以便后续判断使用
                        result['ct_values'][ch_name] = ct_value
                
                if ct_value is not None:
                    item = QTableWidgetItem(f"{ct_value:.2f}")
                    # 根据是否阳性设置颜色（只判断当前项目配置中包含的通道）
                    if ch_name in project_config:
                        ch_config = project_config[ch_name]
                        threshold = ch_config.get('threshold', None)
                        if threshold is not None and ct_value < threshold:
                            item.setBackground(QColor(255, 200, 200))  # 浅红色 - 阳性
                        else:
                            item.setBackground(QColor(200, 255, 200))  # 浅绿色 - 阴性
                    # 如果没有项目配置，不设置背景色（显示默认颜色）
                    self.judgment_table.setItem(row_idx, col_idx, item)
                else:
                    self.judgment_table.setItem(row_idx, col_idx, QTableWidgetItem("N/A"))
            
            # 判读结果（阳性targets）
            result_col_idx = 4 + len(valid_channels)  # 判读结果列索引
            # 过滤掉无效的target（如"\"、"/"等）
            valid_positive_targets = [t for t in result['positive_targets'] 
                                     if t and t.strip() and t.strip() != '\\' and t.strip() != '/']
            if valid_positive_targets:
                result_text = ", ".join(valid_positive_targets)
                item = QTableWidgetItem(result_text)
                item.setBackground(QColor(255, 200, 200))  # 浅红色
                self.judgment_table.setItem(row_idx, result_col_idx, item)
            else:
                self.judgment_table.setItem(row_idx, result_col_idx, QTableWidgetItem("阴性"))
        
        # 调整列宽
        self.judgment_table.resizeColumnsToContents()
    
    def export_judgment_results(self):
        """导出判读结果到Excel文件"""
        # 检查表格是否有数据
        if self.judgment_table.rowCount() == 0:
            QMessageBox.warning(self, '警告', '没有可导出的数据')
            return
        
        # 使用文件对话框让用户选择保存位置（默认xlsx格式）
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            '导出判读结果', 
            '判读结果.xlsx', 
            'Excel Files (*.xlsx);;All Files (*)'
        )
        
        if not file_path:
            return  # 用户取消了保存
        
        # 确保文件扩展名正确
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'
        
        try:
            # 提取表格数据
            data = []
            headers = []
            
            # 获取表头
            for col in range(self.judgment_table.columnCount()):
                header_item = self.judgment_table.horizontalHeaderItem(col)
                if header_item:
                    headers.append(header_item.text())
                else:
                    headers.append(f'列{col+1}')
            
            # 获取表格数据
            for row in range(self.judgment_table.rowCount()):
                row_data = []
                for col in range(self.judgment_table.columnCount()):
                    item = self.judgment_table.item(row, col)
                    if item:
                        row_data.append(item.text())
                    else:
                        row_data.append('')
                data.append(row_data)
            
            # 创建DataFrame
            df = pd.DataFrame(data, columns=headers)
            
            # 使用openpyxl导出xlsx格式
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            QMessageBox.information(self, '成功', f'判读结果已成功导出到:\n{file_path}')
            
        except Exception as e:
            QMessageBox.critical(self, '错误', f'导出失败:\n{str(e)}')


def main():
    app = QApplication(sys.argv)
    window = PCRAnalyzerApp()
    window.showMaximized()  # 启动时最大化窗口
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
