#!/usr/bin/env python3
"""
Word to Markdown Converter

高质量的Word文档转Markdown转换器
- 使用 Docling 的 REFERENCED 模式确保图片保持在原始位置
- 支持从URL下载Word文档
- 自动提取图片并上传到MinIO对象存储
- 生成包含可访问图片链接的Markdown

Author: GitHub Copilot & Joe
Date: 2025-09-03
Version: 1.0.0
"""

import os
import tempfile
import requests
import uuid
import re
from minio import Minio
from docling.document_converter import DocumentConverter
from docling_core.types.doc.base import ImageRefMode
from pathlib import Path

# MinIO 配置
MINIO_ENDPOINT = "127.0.0.1:9000"
MINIO_ACCESS_KEY = "minioadmin"
MINIO_SECRET_KEY = "minioadmin"
MINIO_BUCKET = "img"
MINIO_BASE_URL = "http://127.0.0.1:9000"

def upload_file_to_minio(file_path, object_name):
    """上传文件到 MinIO 对象存储"""
    try:
        client = Minio(
            MINIO_ENDPOINT,
            access_key=MINIO_ACCESS_KEY,
            secret_key=MINIO_SECRET_KEY,
            secure=False
        )
        
        if not client.bucket_exists(MINIO_BUCKET):
            client.make_bucket(MINIO_BUCKET)
        
        unique_filename = f"{uuid.uuid4()}_{object_name}"
        
        client.fput_object(
            MINIO_BUCKET,
            unique_filename,
            file_path,
            content_type="image/png"
        )
        
        return f"{MINIO_BASE_URL}/{MINIO_BUCKET}/{unique_filename}"
        
    except Exception as e:
        print(f"⚠️ 文件上传失败: {e}")
        return None

def process_referenced_images(artifacts_dir, markdown_content):
    """处理REFERENCED模式生成的图片文件，上传到MinIO并更新链接"""
    
    if not artifacts_dir or not artifacts_dir.exists():
        return markdown_content
    
    image_files = list(artifacts_dir.glob("*.png"))
    
    if not image_files:
        print("📸 没有找到图片文件")
        return markdown_content
    
    print(f"📸 找到 {len(image_files)} 张图片，正在上传到 MinIO...")
    
    for image_file in image_files:
        try:
            minio_url = upload_file_to_minio(str(image_file), image_file.name)
            
            if minio_url:
                pattern = rf'!\[[^\]]*\]\([^)]*{re.escape(image_file.name)}[^)]*\)'
                new_image_ref = f'![Image]({minio_url})'
                markdown_content = re.sub(pattern, new_image_ref, markdown_content)
                
                print(f"✅ 图片已上传并替换: {image_file.name} -> {minio_url}")
            
        except Exception as e:
            print(f"⚠️ 处理图片 {image_file.name} 失败: {e}")
    
    return markdown_content

def word_to_markdown_from_url(url, use_referenced_mode=True):
    """从 URL 下载 Word 文档并转换为 Markdown"""
    
    converter = DocumentConverter()
    
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        print("📥 正在下载Word文档...")
        response = requests.get(url, stream=True)
        response.raise_for_status()
        
        temp_file = temp_path / "document.docx"
        with open(temp_file, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        print("🔄 正在转换文档...")
        result = converter.convert(temp_file)
        
        if use_referenced_mode:
            output_dir = temp_path / "output"
            artifacts_dir = output_dir / "input_artifacts"
            output_dir.mkdir(exist_ok=True)
            artifacts_dir.mkdir(parents=True, exist_ok=True)
            
            markdown_file = output_dir / "input.md"
            result.document.save_as_markdown(
                filename=markdown_file,
                artifacts_dir=artifacts_dir,
                image_mode=ImageRefMode.REFERENCED
            )
            
            with open(markdown_file, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            print("🖼️ 正在处理图片...")
            markdown_content = process_referenced_images(artifacts_dir, markdown_content)
            
        else:
            markdown_content = result.document.export_to_markdown()
            print("⚠️ 使用传统模式，图片可能不在原始位置")
        
        return markdown_content

def main():
    """主测试函数"""
    test_url = "http://127.0.0.1:9000/img/%E4%B8%AD%E9%93%B6%E8%AF%81%E5%88%B8-%E5%A4%A7%E6%A8%A1%E5%9E%8B%E8%AF%86%E9%97%AE%E7%AD%94POC%E6%96%B9%E6%A1%880326-v1.4.docx?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=T9WJS99TB65I1BF8ILHR%2F20250903%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20250903T124900Z&X-Amz-Expires=604800&X-Amz-Security-Token=eyJhbGciOiJIUzUxMiIsInR5cCI6IkpXVCJ9.eyJhY2Nlc3NLZXkiOiJUOVdKUzk5VEI2NUkxQkY4SUxIUiIsImV4cCI6MTc1Njk0NjUzNywicGFyZW50IjoibWluaW9hZG1pbiJ9.y0Sv8Gjeh5N5iOOBG8eqBoAAV2nW-yYLSCYCiGT6eRtTWnBBDtqG3P8Rn7Szo8wzaiPuWJ7XCatQlg4BhdYlrQ&X-Amz-SignedHeaders=host&versionId=null&X-Amz-Signature=4be5d400b22342cc31d392c2b28eda1942f46df92e395efa095a00dfca9e8085"
    
    print("🚀 Word转Markdown转换器 - 最终版本")
    print("✨ 特性：")
    print("   • 使用Docling REFERENCED模式保持图片原始位置")
    print("   • 自动上传图片到MinIO对象存储")
    print("   • 支持从URL直接下载Word文档")
    print("   • 生成包含可访问图片链接的Markdown")
    print("-" * 60)
    
    try:
        markdown = word_to_markdown_from_url(test_url, use_referenced_mode=True)
        
        print("\n✅ 转换成功!")
        print(f"📊 结果长度: {len(markdown)} 字符")
        print(f"🖼️ 图片位置: 保持在原始文档位置")
        
        output_file = "output_improved.md"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(markdown)
        
        print(f"💾 完整结果已保存到 {output_file}")
        
        image_count = len(re.findall(r'!\[.*?\]\(http://127\.0\.0\.1:9000/img/.*?\)', markdown))
        print(f"📸 成功处理图片数量: {image_count}")
        
        print("\n📝 前 300 字符预览:")
        print("-" * 50)
        print(markdown[:300])
        print("-" * 50)
        
        if image_count > 0:
            print(f"\n🔗 图片链接示例:")
            for match in re.finditer(r'!\[.*?\]\((http://127\.0\.0\.1:9000/img/.*?)\)', markdown):
                print(f"   {match.group(1)}")
                break
            
    except Exception as e:
        print(f"❌ 转换失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
