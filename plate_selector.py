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
        
        self.create_plate_grid()
        
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
        
        # 添加列标题（1-12或1-24）
        for col in range(1, cols + 1):
            label = QLabel(str(col))
            label.setAlignment(Qt.AlignCenter)
            label.setMinimumWidth(30)
            label.setMaximumWidth(30)
            self.plate_layout.addWidget(label, 0, col)
        
        # 添加行标题和按钮
        row_labels = [chr(65 + i) for i in range(rows)]  # A, B, C, ...
        
        for row_idx, row_label in enumerate(row_labels, 1):
            # 行标签
            label = QLabel(row_label)
            label.setAlignment(Qt.AlignCenter)
            label.setMinimumWidth(25)
            label.setMaximumWidth(25)
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
            # 如果有Ct值，显示在按钮上
            if 'ct' in data and pd.notna(data['ct']):
                btn.setText(f"{data['ct']:.1f}")
                btn.setToolTip(f"孔位 {well_name}\nCt值: {data['ct']:.2f}")
            else:
                btn.setText("")
                btn.setToolTip(f"孔位 {well_name}")
            
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


# 修复导入
import pandas as pd

