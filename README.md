# 🧠 智能测试用例生成工具 (AI TestCase Generator)

<div align="center">

基于人工智能的自动化测试用例生成系统，支持多种文档格式，集成大语言模型，自动生成高质量的功能测试和接口测试用例。

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/guoguangning/TestCaseGeneration?style=social)](https://github.com/guoguangning/TestCaseGeneration)

[快速开始](#快速开始) • [使用方法](#使用方法) • [常见问题](#常见问题) • [贡献指南](#贡献指南)

</div>

---

## ✨ 核心价值

- 🚀 **高效生成**：AI 驱动，将测试用例生成时间从数小时缩短至数分钟
- 📚 **多格式支持**：支持 7+ 种文档格式，无需手动转换
- 🎯 **专业方法**：集成 8 种经典测试用例设计方法
- 🏭 **行业优化**：针对 12+ 个行业领域定制优化
- 🔄 **批量处理**：一次性处理多个需求文档

## 🎯 适用场景

- 软件测试团队快速生成测试用例
- 需求评审阶段的用例设计
- 回归测试用例补充
- 接口自动化测试用例生成

---

## 主要功能

### 📄 多格式文档支持

支持读取和处理以下格式的文档：
- ✅ Word 文档 (.docx)
- ✅ PDF 文档 (.pdf)
- ✅ Excel 表格 (.xlsx)
- ✅ Markdown 文件 (.md)
- ✅ 纯文本文件 (.txt)
- ✅ JSON 文件 (.json)
- ✅ YAML 文件 (.yml/.yaml)

### 🧠 AI 智能能力

- 🔍 **智能内容提取**：自动提取文本、表格、图片中的关键内容
- 📷 **OCR 文字识别**：使用 PaddleOCR 对图片中的文字进行识别
- 🧠 **大语言模型集成**：基于 DeepSeek V3.2 模型
- 🎯 **行业知识库优化**：针对不同行业领域的术语和规则优化
- 📝 **智能提示词生成**：根据选择的参数自动生成优化的提示词

### 🧪 多种测试用例设计方法

1. **等价类划分** - 有效/无效等价类测试
2. **边界值分析** - 极值和临界点测试
3. **决策表** - 复杂业务规则全覆盖
4. **状态转换** - 状态机模型测试
5. **错误推测** - 基于经验的异常测试
6. **场景法** - 端到端用户旅程测试
7. **因果图** - 输入条件逻辑分析
8. **正交分析法** - 参数组合优化覆盖

### 🏭 行业特定优化

针对以下行业进行优化：
- 互联网/电子商务
- 保险业
- 金融科技
- 医疗健康
- 教育科技
- 游戏开发
- 物联网
- 人工智能
- 大数据
- 云计算
- 汽车电子

### 📤 多种导出格式

支持将生成的测试用例导出为：
- Word 文档 (.docx)
- Text 文件 (.txt)
- Markdown (.md)
- JSON（结构化数据）
- Excel（可编辑表格）

---

## 系统要求

### 硬件要求
- **操作系统**：Windows 10/11, macOS 10.14+, Linux（Ubuntu 18.04+）
- **内存**：建议 8GB+（处理大型文档时需要更多）
- **存储空间**：至少 2GB 可用空间（用于依赖和模型）
- **网络**：稳定的网络连接（调用 API 和下载 OCR 模型）

### 软件要求
- **Python**：3.8 或更高版本
- **pip**：最新版本（建议 21.0+）
- **Git**：克隆代码需要（可选）

### 依赖要求
- PyQt5（GUI 界面）
- OpenAI SDK（AI 模型调用）
- PaddleOCR & PaddlePaddle（OCR 识别）
- 其他依赖见 requirements.txt

---

## 快速开始

### 📦 方法一：直接运行（推荐）

```bash
# 1. 克隆仓库
git clone https://github.com/guoguangning/TestCaseGeneration.git
cd TestCaseGeneration

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置 API（编辑 api_config.json）
# 修改 api_key, base_url, model

# 4. 运行程序
python ai_assistant_tester_v1.1.0.py
```

### 📦 方法二：使用虚拟环境（推荐给开发者）

```bash
# 1. 创建虚拟环境
python -m venv venv

# 2. 激活虚拟环境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行程序
python ai_assistant_tester_v1.1.0.py
```

---

## 使用方法

### 🔧 第一步：配置 API

1. 打开 `api_config.json` 文件
2. 填写您的 API 信息：
   ```json
   {
     "api_key": "your-api-key-here",
     "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
     "model": "deepseek-v3.2"
   }
   ```
3. 保存文件

### 📂 第二步：准备需求文档

将您的需求文档放入一个文件夹，支持的格式：
- Word 文档（.docx）
- PDF 文档（.pdf）
- Excel 表格（.xlsx）
- Markdown 文件（.md）
- 纯文本文件（.txt）
- JSON 文件（.json）
- YAML 文件（.yml/.yaml）

### 🚀 第三步：生成测试用例

#### 3.1 添加知识库
1. 点击"添加知识库"按钮
2. 选择包含需求文档的文件夹
3. 文档会自动加载到文件列表中

#### 3.2 选择文档
- **全选**：点击"全选"按钮选中所有文档
- **单选/多选**：在文件列表中点击选择

#### 3.3 配置参数
- **行业选择**：根据您的项目类型选择行业
- **提示词模式**：
  - `文档`：使用文档内容作为上下文
  - `参数输入`：手动输入参数进行测试用例设计
- **功能模式**：
  - `功能测试用例`：生成功能测试用例
  - `接口测试用例`：生成接口测试用例
- **用例设计方法**：
  - 可多选，支持 8 种经典方法
  - 系统会自动生成优化的提示词

#### 3.4 生成用例
1. 点击"开始推理"按钮
2. 等待 AI 生成测试用例（可能需要几秒钟到几分钟）
3. 生成的结果会显示在"生成结果"区域

#### 3.5 导出结果
1. 选择导出格式（Word/Text/Markdown/JSON/XLSX）
2. 点击"导出结果"按钮
3. 选择保存路径
4. 完成！

### 💡 高级功能

#### 自定义标题提取
- **文本标题**：指定需要提取的章节标题（用逗号分隔）
- **表格标题**：指定需要提取的表格标题
- **图片标题**：指定需要提取的图片标题（会自动进行 OCR 识别）

#### 提示词优化
- 点击"更新提示词"按钮
- 系统会根据您选择的参数自动生成优化的提示词
- 您也可以手动编辑提示词进行微调

---

## 配置说明

### API 配置（api_config.json）

```json
{
  "api_key": "您的 API 密钥",
  "base_url": "API 基础 URL",
  "model": "模型名称"
}
```

### 可配置参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| api_key | API 访问密钥 | - |
| base_url | API 服务地址 | https://dashscope.aliyuncs.com/compatible-mode/v1 |
| model | 使用的模型名称 | deepseek-v3.2 |

---

## 项目结构

```
TestCaseGeneration/
├── ai_assistant_tester_v1.1.0.py    # 主程序（GUI应用）
├── ultimate_json_to_excel_v3.py      # Markdown转Excel工具
├── main.py                           # PyCharm模板文件
├── requirements.txt                  # Python依赖包
├── api_config.json                   # API配置文件
└── README.md                         # 项目说明文档
```

---

## 技术架构

### 核心技术栈

- **GUI 框架**：PyQt5
- **AI 模型**：DeepSeek V3.2（通过阿里云 DashScope）
- **OCR 引擎**：PaddleOCR + PaddlePaddle
- **文档处理**：python-docx, PyPDF2, openpyxl, PyYAML
- **数据处理**：pandas, numpy
- **图像处理**：opencv-python

### 系统架构

```
┌─────────────────────────────────────────┐
│         PyQt5 GUI 界面层                │
├─────────────────────────────────────────┤
│         业务逻辑层                       │
│  - 文档解析  - 内容提取  - OCR识别      │
├─────────────────────────────────────────┤
│         AI 调用层                        │
│  - 提示词生成  - API调用  - 结果处理    │
├─────────────────────────────────────────┤
│         数据处理层                       │
│  - JSON处理  - 数据转换  - 格式化      │
├─────────────────────────────────────────┤
│         文件操作层                       │
│  - 多格式读取  - 多格式导出             │
└─────────────────────────────────────────┘
```

---

## 依赖项

项目依赖以下 Python 库：

- paddleocr==2.10.0 - OCR 文字识别
- paddlepaddle==3.0.0 - PaddleOCR 的依赖
- pandas==2.2.3 - 数据处理
- PyPDF2==3.0.1 - PDF 文件处理
- PyQt5==5.15.11 - 图形用户界面
- python-docx==1.1.2 - Word 文档处理
- openpyxl==3.1.5 - Excel 文件处理
- opencv-python==4.11.0.86 - 图像处理
- Markdown==3.7 - Markdown 文件处理
- lxml==5.3.1 - XML 处理
- ollama==0.4.7 - 本地模型支持
- PyYAML==6.0.2 - YAML 文件处理
- nltk==3.9.1 - 自然语言处理
- openai==1.65.1 - OpenAI API 客户端
- numpy~=2.2.4 - 数值计算
- tabulate==0.9.0 - 表格格式化

安装命令：
```bash
pip install -r requirements.txt
```

---

## 常见问题（FAQ）

### 安装相关

**Q1: 安装 paddlepaddle 时报错怎么办？**
A: 请确保 Python 版本为 3.8+，并使用虚拟环境安装。如果仍然失败，请检查网络连接或尝试使用国内镜像源。

**Q2: 如何查看已安装的包版本？**
A: 运行 `pip list` 命令查看所有已安装的包及其版本。

**Q3: 首次运行时下载 OCR 模型很慢怎么办？**
A: 首次使用 PaddleOCR 时会自动下载模型，请耐心等待。也可以手动下载模型并指定本地路径。

### 运行相关

**Q4: 点击"开始推理"后没有反应？**
A: 请检查：
1. API 配置是否正确
2. 网络连接是否正常
3. 是否选择了文档
4. 查看终端输出的错误信息

**Q5: OCR 识别失败怎么办？**
A: 
1. 首次使用时会自动下载 OCR 模型，请确保网络通畅
2. 检查图片格式是否支持（JPG、PNG等）
3. 查看终端输出的详细错误信息

**Q6: 生成的测试用例质量不高怎么办？**
A: 
1. 提供清晰、详细的需求文档
2. 选择合适的行业领域
3. 尝试不同的测试用例设计方法组合
4. 手动编辑提示词进行微调

### 功能相关

**Q7: 可以同时处理多个文档吗？**
A: 可以！使用"全选"按钮或按住 Ctrl/Cmd 键多选文档。

**Q8: 如何批量转换 Markdown 到 Excel？**
A: 使用 `ultimate_json_to_excel_v3.py` 工具，修改配置中的输入输出路径即可。

### 性能相关

**Q9: 处理大型文档时很慢怎么办？**
A:
1. 系统会自动截断过长的上下文（15000字符）
2. 建议分批处理大型文档
3. 关闭其他占用内存的程序

**Q10: 程序占用内存过大？**
A: 
1. 关闭不需要的文档预览
2. 处理完一个文档后刷新列表
3. 增加系统虚拟内存

---

## 注意事项

1. **API 密钥安全**：请妥善保管您的 API 密钥，不要将其提交到公开代码仓库
2. **首次使用**：首次使用 OCR 功能时会下载模型，请确保网络通畅
3. **文档格式**：建议使用格式规范的需求文档，以获得更好的生成效果
4. **上下文长度**：系统会自动截断过长的上下文（15000字符），确保关键信息在前部
5. **网络连接**：调用 AI 模型需要稳定的网络连接
6. **批量处理**：处理大量文档时建议分批进行，避免内存溢出

---

## 更新日志

查看详细的更新历史，请访问 [CHANGELOG.md](CHANGELOG.md)（待创建）

### 当前版本 v1.1.0

- ✨ 新增接口测试用例生成功能
- 🐛 修复 JSON 格式转换问题
- 🔧 优化 OCR 识别准确率
- 📝 完善用户界面交互
- 🚀 支持更多文档格式
- 📚 完善文档说明

---

## 贡献指南

我们欢迎任何形式的贡献！

### 如何贡献

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

详细贡献指南请查看 [CONTRIBUTING.md](CONTRIBUTING.md)（待创建）

---

## 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

---

## 联系方式

- **作者**：郭光宁
- **GitHub**：[@guoguangning](https://github.com/guoguangning)
- **项目地址**：https://github.com/guoguangning/TestCaseGeneration
- **问题反馈**：[提交 Issue](https://github.com/guoguangning/TestCaseGeneration/issues)

---

## 致谢

感谢以下开源项目：

- [DeepSeek](https://www.deepseek.com/) - 提供强大的 AI 模型
- [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) - 优秀的 OCR 引擎
- [PyQt5](https://www.riverbankcomputing.com/software/pyqt/) - 强大的 GUI 框架
- [OpenAI](https://openai.com/) - 提供 API 客户端

---

## 相关链接

- [安装指南](docs/INSTALL.md)（待创建）
- [用户手册](docs/USER_GUIDE.md)（待创建）
- [开发者文档](docs/DEVELOPER.md)（待创建）
- [故障排除](docs/TROUBLESHOOTING.md)（待创建）
- [API 文档](docs/API.md)（待创建）

---

<div align="center">

**如果这个项目对您有帮助，请给一个 ⭐️ Star**

Made with ❤️ by [Guo Guangning](https://github.com/guoguangning)

</div>
