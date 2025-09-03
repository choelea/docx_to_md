# Word to Markdown Converter

一个高质量的Word文档转Markdown转换器，使用Docling库实现，支持图片处理和MinIO对象存储。

## ✨ 特性

- 🔄 **高质量转换** - 基于Docling库的精准Word转Markdown
- 🖼️ **智能图片处理** - 保持图片在原始文档中的位置
- ☁️ **云存储支持** - 自动上传图片到MinIO对象存储
- 🌐 **URL下载** - 支持从URL直接下载Word文档
- 📝 **格式保持** - 保留文档的结构和格式

## 🚀 快速开始

### 环境要求

- Python 3.10+
- conda环境（推荐使用`3.10_docling_api`）

### 安装依赖

```bash
pip install -r requirements.txt
```

### 配置MinIO

在使用前，请确保MinIO服务可用，并在代码中配置相关参数：

```python
# MinIO配置 (在word_to_markdown.py中修改)
MINIO_ENDPOINT = "your-minio-endpoint"
MINIO_ACCESS_KEY = "your-access-key"
MINIO_SECRET_KEY = "your-secret-key"
MINIO_BUCKET_NAME = "your-bucket-name"
```

### 使用方法

```python
from word_to_markdown import word_to_markdown_from_url

# 从URL转换Word文档
url = "https://example.com/document.docx"
markdown_text = word_to_markdown_from_url(url)
print(markdown_text)
```

## 📁 项目结构

```
docx_to_md/
├── word_to_markdown.py    # 核心转换程序
├── requirements.txt       # Python依赖包
├── input.docx            # 测试用Word文档
└── README.md             # 项目说明
```

## 🔧 技术实现

### 核心技术栈

- **Docling 2.5.0** - 文档转换引擎
- **MinIO 7.2.16** - 对象存储客户端
- **PIL/Pillow** - 图片处理
- **Requests** - HTTP客户端

### 关键特性

1. **REFERENCED模式** - 使用Docling的`ImageRefMode.REFERENCED`确保图片位置准确
2. **图片处理流水线** - 提取→转换→上传→链接替换
3. **错误处理** - 完善的异常处理和日志记录

## 📝 API说明

### 主要函数

#### `word_to_markdown_from_url(url)`

从URL下载Word文档并转换为Markdown。

**参数:**
- `url` (str): Word文档的URL地址

**返回:**
- `str`: 转换后的Markdown文本

#### `process_referenced_images(doc_result, temp_dir)`

处理文档中的图片引用。

**参数:**
- `doc_result`: Docling转换结果对象
- `temp_dir` (str): 临时目录路径

**返回:**
- `dict`: 图片ID到MinIO URL的映射

## 🧪 测试

项目包含测试用的Word文档`input.docx`，可以用于验证转换功能：

```python
# 测试本地文件转换
with open("input.docx", "rb") as f:
    # 实现本地文件转换逻辑
    pass
```

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目。

## 📞 支持

如有问题，请在GitHub上创建Issue。
