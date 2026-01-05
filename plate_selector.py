"""
孔板选择器组件
支持96孔板和384孔板的可视化选择
"""
from PyQt5.QtWidgets import (QWidget, QGridLayout, QPushButton, QLabel, 
                             QGroupBox, QVBoxLayout, QHBoxLayout, QComboBox)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor, QPalette
import pandas as pd


class PlateSelector(QWidget):
    """孔板选择器"""
    
    # 信号：当孔被选中时发出
    well_selected = pyqtSignal(str)  # 发出孔位名称，如"A1", "B2"
    
    def __init__(self, plate_type='96', parent=None):
        super().__init__(parent)
        self.plate_type = plate_type  # '96' 或 '384'
        self.selected_wells = set()  # 选中的孔位集合
        self.well_buttons = {}  # 存储按钮引用
        self.well_data = {}  # 存储每个孔的数据
        self.init_ui()
    
    def init_ui(self):
        """初始化界面"""
        layout = QVBoxLayout(self)
        
        # 标题
        title = QLabel(f'{self.plate_type}孔板')
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(title)
        
        # 孔板网格
        self.plate_widget = QWidget()
        self.plate_layout = QGridLayout(self.plate_widget)
        self.plate_layout.setSpacing(2)
        self.plate_layout.setContentsMargins(5, 5, 5, 5)  # 设置边距
        
        self.create_plate_grid()
        
        # 设置plate_widget的尺寸策略，确保内容不被裁剪
        from PyQt5.QtWidgets import QSizePolicy
        self.plate_widget.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.MinimumExpanding)
        
        layout.addWidget(self.plate_widget)
        layout.addStretch()
    
    def create_plate_grid(self):
        """创建孔板网格"""
        if self.plate_type == '96':
            rows = 8  # A-H
            cols = 12  # 1-12
        else:  # 384
            rows = 16  # A-P
            cols = 24  # 1-24
        
        # 添加列标题（1-12或1-24），可点击全选列
        self.col_labels = {}
        for col in range(1, cols + 1):
            label = QPushButton(str(col))
            label.setMinimumWidth(30)
            label.setMaximumWidth(30)
            label.setMinimumHeight(25)
            label.setMaximumHeight(25)
            label.setStyleSheet("""
                QPushButton {
                    background-color: #e0e0e0;
                    border: 1px solid #999;
                    border-radius: 2px;
                    font-weight: bold;
                    text-align: center;
                }
                QPushButton:hover {
                    background-color: #d0d0d0;
                }
            """)
            # 连接点击事件，全选该列
            label.clicked.connect(lambda checked, c=col: self.select_column(c))
            self.col_labels[col] = label
            self.plate_layout.addWidget(label, 0, col)
        
        # 添加行标题和按钮
        row_labels = [chr(65 + i) for i in range(rows)]  # A, B, C, ...
        
        # 存储行标签引用
        self.row_labels = {}
        for row_idx, row_label in enumerate(row_labels, 1):
            # 行标签，可点击全选行
            label = QPushButton(row_label)
            label.setMinimumWidth(25)
            label.setMaximumWidth(25)
            label.setMinimumHeight(30)
            label.setMaximumHeight(30)
            label.setStyleSheet("""
                QPushButton {
                    background-color: #e0e0e0;
                    border: 1px solid #999;
                    border-radius: 2px;
                    font-weight: bold;
                    text-align: center;
                }
                QPushButton:hover {
                    background-color: #d0d0d0;
                }
            """)
            # 连接点击事件，全选该行
            label.clicked.connect(lambda checked, r=row_label: self.select_row(r))
            self.row_labels[row_label] = label
            self.plate_layout.addWidget(label, row_idx, 0)
            
            # 创建按钮
            for col in range(1, cols + 1):
                well_name = f"{row_label}{col}"
                btn = QPushButton("")
                btn.setMinimumSize(30, 30)
                btn.setMaximumSize(30, 30)
                btn.setCheckable(True)
                btn.setStyleSheet(self.get_default_button_style())
                
                # 连接信号
                btn.clicked.connect(lambda checked, name=well_name: self.on_well_clicked(name, checked))
                
                self.well_buttons[well_name] = btn
                self.plate_layout.addWidget(btn, row_idx, col)
    
    def get_default_button_style(self):
        """获取默认按钮样式"""
        return """
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
                border: 1px solid #999;
            }
            QPushButton:checked {
                background-color: #4CAF50;
                border: 2px solid #2E7D32;
            }
        """
    
    def on_well_clicked(self, well_name, checked):
        """孔位被点击时的处理"""
        if checked:
            self.selected_wells.add(well_name)
            # 更新按钮样式
            self.well_buttons[well_name].setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 2px solid #2E7D32;
                    border-radius: 3px;
                }
            """)
        else:
            self.selected_wells.discard(well_name)
            # 恢复默认样式
            if well_name in self.well_data:
                self.update_well_style(well_name, self.well_data[well_name])
            else:
                self.well_buttons[well_name].setStyleSheet(self.get_default_button_style())
        
        # 发出信号
        self.well_selected.emit(well_name)
    
    def clear_selection(self):
        """清除所有选择"""
        # 先清除选中集合
        self.selected_wells.clear()
        
        # 更新所有按钮的状态和样式
        for well_name, btn in self.well_buttons.items():
            # 临时断开信号，避免触发on_well_clicked
            btn.blockSignals(True)
            btn.setChecked(False)
            btn.blockSignals(False)
            
            # 恢复按钮样式
            if well_name in self.well_data:
                self.update_well_style(well_name, self.well_data[well_name])
            else:
                btn.setStyleSheet(self.get_default_button_style())
        
        # 发出信号通知选择已清除（使用空字符串）
        self.well_selected.emit("")
    
    def set_well_data(self, well_name, data):
        """设置孔位数据（用于显示Ct值等信息）"""
        self.well_data[well_name] = data
        
        # 更新按钮显示
        if well_name in self.well_buttons:
            btn = self.well_buttons[well_name]
            # 如果有Ct值，显示在按钮上（显示为整数）
            if 'ct' in data and pd.notna(data['ct']):
                ct_value = data['ct']
                # 转换为整数显示
                ct_int = int(round(ct_value))
                btn.setText(str(ct_int))
                btn.setToolTip(f"孔位 {well_name}\nCt值: {ct_value:.2f}")
            else:
                # 没有CT值时显示N/A
                btn.setText("N/A")
                btn.setToolTip(f"孔位 {well_name}\nCt值: 无")
            
            # 根据数据状态更新颜色
            self.update_well_style(well_name, data)
    
    def update_well_style(self, well_name, data):
        """根据数据更新孔位样式"""
        if well_name not in self.well_buttons:
            return
        
        btn = self.well_buttons[well_name]
        
        # 如果孔位被选中，保持选中样式
        if well_name in self.selected_wells:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 2px solid #2E7D32;
                    border-radius: 3px;
                }
            """)
            return
        
        # 根据数据状态设置颜色
        if 'status' in data:
            status = data['status']
            if status == 'positive':
                color = '#FF6B6B'  # 红色 - 阳性
            elif status == 'negative':
                color = '#95E1D3'  # 浅绿 - 阴性
            elif status == 'invalid':
                color = '#FFD93D'  # 黄色 - 无效
            else:
                color = '#f0f0f0'  # 灰色 - 无数据
        else:
            # 根据Ct值判断
            if 'ct' in data and pd.notna(data['ct']):
                ct = data['ct']
                if ct < 30:
                    color = '#FF6B6B'  # 红色 - 强阳性
                elif ct < 35:
                    color = '#FFA07A'  # 橙色 - 弱阳性
                else:
                    color = '#95E1D3'  # 浅绿 - 阴性
            else:
                color = '#f0f0f0'  # 灰色 - 无数据
        
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                border: 1px solid #ccc;
                border-radius: 3px;
            }}
            QPushButton:hover {{
                background-color: {self.lighten_color(color)};
                border: 1px solid #999;
            }}
        """)
    
    def lighten_color(self, color):
        """使颜色变亮"""
        # 简单的颜色变亮处理
        if color == '#f0f0f0':
            return '#ffffff'
        elif color == '#FF6B6B':
            return '#FF8E8E'
        elif color == '#95E1D3':
            return '#B5F1E3'
        elif color == '#FFA07A':
            return '#FFB89A'
        else:
            return color
    
    def get_selected_wells(self):
        """获取选中的孔位列表"""
        return list(self.selected_wells)
    
    def select_row(self, row_label):
        """全选或取消全选指定行"""
        if self.plate_type == '96':
            cols = 12
        else:  # 384
            cols = 24
        
        # 获取该行的所有孔位
        row_wells = [f"{row_label}{col}" for col in range(1, cols + 1)]
        
        # 检查该行是否已全部选中
        all_selected = all(well in self.selected_wells for well in row_wells)
        
        # 如果全部选中，则取消全选；否则全选
        for well_name in row_wells:
            if well_name in self.well_buttons:
                btn = self.well_buttons[well_name]
                btn.blockSignals(True)
                if all_selected:
                    # 取消选中
                    btn.setChecked(False)
                    self.selected_wells.discard(well_name)
                    # 恢复样式
                    if well_name in self.well_data:
                        self.update_well_style(well_name, self.well_data[well_name])
                    else:
                        btn.setStyleSheet(self.get_default_button_style())
                else:
                    # 选中
                    btn.setChecked(True)
                    self.selected_wells.add(well_name)
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: #4CAF50;
                            border: 2px solid #2E7D32;
                            border-radius: 3px;
                        }
                    """)
                btn.blockSignals(False)
        
        # 发出信号
        self.well_selected.emit("")
    
    def select_column(self, col):
        """全选或取消全选指定列"""
        if self.plate_type == '96':
            rows = 8  # A-H
        else:  # 384
            rows = 16  # A-P
        
        row_labels = [chr(65 + i) for i in range(rows)]  # A, B, C, ...
        
        # 获取该列的所有孔位
        col_wells = [f"{row_label}{col}" for row_label in row_labels]
        
        # 检查该列是否已全部选中
        all_selected = all(well in self.selected_wells for well in col_wells)
        
        # 如果全部选中，则取消全选；否则全选
        for well_name in col_wells:
            if well_name in self.well_buttons:
                btn = self.well_buttons[well_name]
                btn.blockSignals(True)
                if all_selected:
                    # 取消选中
                    btn.setChecked(False)
                    self.selected_wells.discard(well_name)
                    # 恢复样式
                    if well_name in self.well_data:
                        self.update_well_style(well_name, self.well_data[well_name])
                    else:
                        btn.setStyleSheet(self.get_default_button_style())
                else:
                    # 选中
                    btn.setChecked(True)
                    self.selected_wells.add(well_name)
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: #4CAF50;
                            border: 2px solid #2E7D32;
                            border-radius: 3px;
                        }
                    """)
                btn.blockSignals(False)
        
        # 发出信号
        self.well_selected.emit("")
    
    def toggle_select_all(self):
        """全选或取消全选所有孔位"""
        # 检查是否已全部选中
        all_selected = len(self.selected_wells) == len(self.well_buttons)
        
        for well_name, btn in self.well_buttons.items():
            btn.blockSignals(True)
            if all_selected:
                # 取消全选
                btn.setChecked(False)
                self.selected_wells.discard(well_name)
                # 恢复样式
                if well_name in self.well_data:
                    self.update_well_style(well_name, self.well_data[well_name])
                else:
                    btn.setStyleSheet(self.get_default_button_style())
            else:
                # 全选
                btn.setChecked(True)
                self.selected_wells.add(well_name)
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #4CAF50;
                        border: 2px solid #2E7D32;
                        border-radius: 3px;
                    }
                """)
            btn.blockSignals(False)
        
        # 发出信号
        self.well_selected.emit("")


# 修复导入
import pandas as pd

