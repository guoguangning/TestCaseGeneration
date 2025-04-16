import json
import os
from typing import Dict, List, Optional
from datetime import datetime

class TemplateManager:
    """测试用例模板管理类"""
    
    def __init__(self, config_path: str = "template_config.json"):
        """
        初始化模板管理器
        
        Args:
            config_path: 模板配置文件路径
        """
        self.config_path = config_path
        self.config = self._load_config()
        self.templates_dir = "templates"
        self._ensure_templates_dir()
        
    def _load_config(self) -> Dict:
        """加载配置文件"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return {"version": "1.0.0", "categories": [], "templates": []}
            
    def _save_config(self):
        """保存配置文件"""
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)
            
    def _ensure_templates_dir(self):
        """确保模板目录存在"""
        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir)
            
    def get_categories(self) -> List[Dict]:
        """获取所有模板分类"""
        return self.config.get("categories", [])
        
    def get_templates(self, category_id: Optional[str] = None) -> List[Dict]:
        """获取模板列表
        
        Args:
            category_id: 分类ID，如果为None则返回所有模板
        """
        templates = self.config.get("templates", [])
        if category_id:
            return [t for t in templates if t["category"] == category_id]
        return templates
        
    def get_template(self, template_id: str) -> Optional[Dict]:
        """获取指定模板
        
        Args:
            template_id: 模板ID
        """
        templates = self.get_templates()
        for template in templates:
            if template["id"] == template_id:
                return template
        return None
        
    def add_template(self, template: Dict) -> bool:
        """添加新模板
        
        Args:
            template: 模板数据
        """
        # 验证模板数据
        if not self._validate_template(template):
            return False
            
        # 检查模板ID是否已存在
        if self.get_template(template["id"]):
            return False
            
        # 添加模板
        self.config["templates"].append(template)
        self._save_config()
        return True
        
    def update_template(self, template_id: str, template_data: Dict) -> bool:
        """更新模板
        
        Args:
            template_id: 模板ID
            template_data: 新的模板数据
        """
        # 验证模板数据
        if not self._validate_template(template_data):
            return False
            
        # 更新模板
        templates = self.config["templates"]
        for i, template in enumerate(templates):
            if template["id"] == template_id:
                templates[i] = template_data
                self._save_config()
                return True
        return False
        
    def delete_template(self, template_id: str) -> bool:
        """删除模板
        
        Args:
            template_id: 模板ID
        """
        templates = self.config["templates"]
        for i, template in enumerate(templates):
            if template["id"] == template_id:
                del templates[i]
                self._save_config()
                return True
        return False
        
    def _validate_template(self, template: Dict) -> bool:
        """验证模板数据格式
        
        Args:
            template: 模板数据
        """
        required_fields = ["id", "name", "category", "version", "structure"]
        return all(field in template for field in required_fields)
        
    def export_template(self, template_id: str, export_path: str) -> bool:
        """导出模板
        
        Args:
            template_id: 模板ID
            export_path: 导出路径
        """
        template = self.get_template(template_id)
        if not template:
            return False
            
        try:
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, ensure_ascii=False, indent=4)
            return True
        except Exception:
            return False
            
    def import_template(self, import_path: str) -> bool:
        """导入模板
        
        Args:
            import_path: 导入文件路径
        """
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                template = json.load(f)
            return self.add_template(template)
        except Exception:
            return False
            
    def create_test_case(self, template_id: str, data: Dict) -> Optional[Dict]:
        """根据模板创建测试用例
        
        Args:
            template_id: 模板ID
            data: 测试用例数据
        """
        template = self.get_template(template_id)
        if not template:
            return None
            
        # 创建测试用例
        test_case = {
            "template_id": template_id,
            "created_at": datetime.now().isoformat(),
            "data": data
        }
        
        # 验证数据是否符合模板结构
        if not self._validate_test_case_data(template["structure"], data):
            return None
            
        return test_case
        
    def _validate_test_case_data(self, structure: Dict, data: Dict) -> bool:
        """验证测试用例数据是否符合模板结构
        
        Args:
            structure: 模板结构
            data: 测试用例数据
        """
        for key, value_type in structure.items():
            if key not in data:
                return False
            if isinstance(value_type, list):
                if not isinstance(data[key], list):
                    return False
            elif isinstance(value_type, dict):
                if not isinstance(data[key], dict):
                    return False
        return True 