import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QComboBox, QPushButton, 
                           QTableWidget, QTableWidgetItem, QMessageBox,
                           QFileDialog, QLineEdit, QTextEdit)
from PyQt5.QtCore import Qt
from template_manager import TemplateManager

class TemplateManagerGUI(QMainWindow):
    """测试用例模板管理GUI"""
    
    def __init__(self):
        super().__init__()
        self.template_manager = TemplateManager()
        self.init_ui()
        
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("测试用例模板管理")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建顶部工具栏
        toolbar = QHBoxLayout()
        
        # 分类选择
        category_label = QLabel("分类:")
        self.category_combo = QComboBox()
        self.update_category_combo()
        toolbar.addWidget(category_label)
        toolbar.addWidget(self.category_combo)
        
        # 模板操作按钮
        self.add_btn = QPushButton("添加模板")
        self.edit_btn = QPushButton("编辑模板")
        self.delete_btn = QPushButton("删除模板")
        self.import_btn = QPushButton("导入模板")
        self.export_btn = QPushButton("导出模板")
        
        toolbar.addWidget(self.add_btn)
        toolbar.addWidget(self.edit_btn)
        toolbar.addWidget(self.delete_btn)
        toolbar.addWidget(self.import_btn)
        toolbar.addWidget(self.export_btn)
        
        main_layout.addLayout(toolbar)
        
        # 创建模板列表
        self.template_table = QTableWidget()
        self.template_table.setColumnCount(5)
        self.template_table.setHorizontalHeaderLabels(["ID", "名称", "分类", "版本", "描述"])
        self.update_template_table()
        main_layout.addWidget(self.template_table)
        
        # 创建模板编辑区域
        edit_layout = QHBoxLayout()
        
        # 左侧：模板基本信息
        basic_info = QVBoxLayout()
        basic_info.addWidget(QLabel("模板ID:"))
        self.template_id_edit = QLineEdit()
        basic_info.addWidget(self.template_id_edit)
        
        basic_info.addWidget(QLabel("模板名称:"))
        self.template_name_edit = QLineEdit()
        basic_info.addWidget(self.template_name_edit)
        
        basic_info.addWidget(QLabel("分类:"))
        self.template_category_combo = QComboBox()
        self.update_category_combo(self.template_category_combo)
        basic_info.addWidget(self.template_category_combo)
        
        basic_info.addWidget(QLabel("版本:"))
        self.template_version_edit = QLineEdit()
        basic_info.addWidget(self.template_version_edit)
        
        edit_layout.addLayout(basic_info)
        
        # 右侧：模板结构
        structure_info = QVBoxLayout()
        structure_info.addWidget(QLabel("模板结构 (JSON):"))
        self.template_structure_edit = QTextEdit()
        structure_info.addWidget(self.template_structure_edit)
        
        edit_layout.addLayout(structure_info)
        
        main_layout.addLayout(edit_layout)
        
        # 底部按钮
        bottom_layout = QHBoxLayout()
        self.save_btn = QPushButton("保存")
        self.cancel_btn = QPushButton("取消")
        bottom_layout.addWidget(self.save_btn)
        bottom_layout.addWidget(self.cancel_btn)
        main_layout.addLayout(bottom_layout)
        
        # 连接信号
        self.connect_signals()
        
    def connect_signals(self):
        """连接信号和槽"""
        self.category_combo.currentTextChanged.connect(self.on_category_changed)
        self.template_table.itemClicked.connect(self.on_template_selected)
        self.add_btn.clicked.connect(self.on_add_template)
        self.edit_btn.clicked.connect(self.on_edit_template)
        self.delete_btn.clicked.connect(self.on_delete_template)
        self.import_btn.clicked.connect(self.on_import_template)
        self.export_btn.clicked.connect(self.on_export_template)
        self.save_btn.clicked.connect(self.on_save_template)
        self.cancel_btn.clicked.connect(self.on_cancel_edit)
        
    def update_category_combo(self, combo_box=None):
        """更新分类下拉框"""
        if combo_box is None:
            combo_box = self.category_combo
            
        combo_box.clear()
        categories = self.template_manager.get_categories()
        for category in categories:
            combo_box.addItem(category["name"], category["id"])
            
    def update_template_table(self):
        """更新模板列表"""
        self.template_table.setRowCount(0)
        category_id = self.category_combo.currentData()
        templates = self.template_manager.get_templates(category_id)
        
        for template in templates:
            row = self.template_table.rowCount()
            self.template_table.insertRow(row)
            self.template_table.setItem(row, 0, QTableWidgetItem(template["id"]))
            self.template_table.setItem(row, 1, QTableWidgetItem(template["name"]))
            self.template_table.setItem(row, 2, QTableWidgetItem(template["category"]))
            self.template_table.setItem(row, 3, QTableWidgetItem(template["version"]))
            self.template_table.setItem(row, 4, QTableWidgetItem(str(template.get("description", ""))))
            
    def on_category_changed(self):
        """分类改变时的处理"""
        self.update_template_table()
        
    def on_template_selected(self, item):
        """选择模板时的处理"""
        row = item.row()
        template_id = self.template_table.item(row, 0).text()
        template = self.template_manager.get_template(template_id)
        
        if template:
            self.template_id_edit.setText(template["id"])
            this.template_name_edit.setText(template["name"])
            index = self.template_category_combo.findData(template["category"])
            if index >= 0:
                this.template_category_combo.setCurrentIndex(index)
            this.template_version_edit.setText(template["version"])
            this.template_structure_edit.setText(str(template["structure"]))
            
    def on_add_template(self):
        """添加模板"""
        self.clear_template_form()
        self.template_id_edit.setEnabled(True)
        self.template_name_edit.setEnabled(True)
        self.template_category_combo.setEnabled(True)
        self.template_version_edit.setEnabled(True)
        self.template_structure_edit.setEnabled(True)
        self.save_btn.setEnabled(True)
        this.cancel_btn.setEnabled(True)
        
    def on_edit_template(self):
        """编辑模板"""
        current_row = this.template_table.currentRow()
        if current_row >= 0:
            template_id = this.template_table.item(current_row, 0).text()
            template = this.template_manager.get_template(template_id)
            if template:
                this.template_id_edit.setText(template["id"])
                this.template_name_edit.setText(template["name"])
                index = this.template_category_combo.findData(template["category"])
                if index >= 0:
                    this.template_category_combo.setCurrentIndex(index)
                this.template_version_edit.setText(template["version"])
                this.template_structure_edit.setText(str(template["structure"]))
                
                this.template_id_edit.setEnabled(False)
                this.template_name_edit.setEnabled(True)
                this.template_category_combo.setEnabled(True)
                this.template_version_edit.setEnabled(True)
                this.template_structure_edit.setEnabled(True)
                this.save_btn.setEnabled(True)
                this.cancel_btn.setEnabled(True)
                
    def on_delete_template(self):
        """删除模板"""
        current_row = this.template_table.currentRow()
        if current_row >= 0:
            template_id = this.template_table.item(current_row, 0).text()
            reply = QMessageBox.question(this, "确认删除", 
                                       f"确定要删除模板 {template_id} 吗？",
                                       QMessageBox.Yes | QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                if this.template_manager.delete_template(template_id):
                    this.update_template_table()
                    QMessageBox.information(this, "成功", "模板已删除")
                else:
                    QMessageBox.warning(this, "错误", "删除模板失败")
                    
    def on_import_template(self):
        """导入模板"""
        file_path, _ = QFileDialog.getOpenFileName(this, "选择模板文件", "", 
                                                 "JSON Files (*.json)")
        if file_path:
            if this.template_manager.import_template(file_path):
                this.update_template_table()
                QMessageBox.information(this, "成功", "模板导入成功")
            else:
                QMessageBox.warning(this, "错误", "模板导入失败")
                
    def on_export_template(self):
        """导出模板"""
        current_row = this.template_table.currentRow()
        if current_row >= 0:
            template_id = this.template_table.item(current_row, 0).text()
            file_path, _ = QFileDialog.getSaveFileName(this, "保存模板文件", 
                                                     f"{template_id}.json",
                                                     "JSON Files (*.json)")
            if file_path:
                if this.template_manager.export_template(template_id, file_path):
                    QMessageBox.information(this, "成功", "模板导出成功")
                else:
                    QMessageBox.warning(this, "错误", "模板导出失败")
                    
    def on_save_template(self):
        """保存模板"""
        template_data = {
            "id": this.template_id_edit.text(),
            "name": this.template_name_edit.text(),
            "category": this.template_category_combo.currentData(),
            "version": this.template_version_edit.text(),
            "structure": eval(this.template_structure_edit.toPlainText())
        }
        
        if this.template_id_edit.isEnabled():
            # 添加新模板
            if this.template_manager.add_template(template_data):
                this.update_template_table()
                QMessageBox.information(this, "成功", "模板添加成功")
            else:
                QMessageBox.warning(this, "错误", "模板添加失败")
        else:
            # 更新现有模板
            if this.template_manager.update_template(template_data["id"], template_data):
                this.update_template_table()
                QMessageBox.information(this, "成功", "模板更新成功")
            else:
                QMessageBox.warning(this, "错误", "模板更新失败")
                
        this.clear_template_form()
        
    def on_cancel_edit(self):
        """取消编辑"""
        this.clear_template_form()
        
    def clear_template_form(self):
        """清空模板表单"""
        this.template_id_edit.clear()
        this.template_name_edit.clear()
        this.template_category_combo.setCurrentIndex(0)
        this.template_version_edit.clear()
        this.template_structure_edit.clear()
        
        this.template_id_edit.setEnabled(False)
        this.template_name_edit.setEnabled(False)
        this.template_category_combo.setEnabled(False)
        this.template_version_edit.setEnabled(False)
        this.template_structure_edit.setEnabled(False)
        this.save_btn.setEnabled(False)
        this.cancel_btn.setEnabled(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TemplateManagerGUI()
    window.show()
    sys.exit(app.exec_()) 