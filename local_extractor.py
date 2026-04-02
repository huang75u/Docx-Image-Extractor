import zipfile
import xml.etree.ElementTree as ET
import os

def main():
    print("="*50)
    print("   本地 Word (.docx) 图片顺序提取工具")
    print("="*50)
    
    file_path = input("请拖入或输入 .docx 文件的路径:\n> ").strip()
    
    # 处理拖拽文件时路径可能带有的双引号
    if file_path.startswith('"') and file_path.endswith('"'):
        file_path = file_path[1:-1]
    
    if not os.path.exists(file_path):
        print(f"\n[错误] 找不到文件: {file_path}")
        os.system("pause")
        return
        
    if not file_path.lower().endswith('.docx'):
        print("\n[错误] 请确保提供的是 .docx 格式的文件！")
        os.system("pause")
        return

    # 在被提取的文档同级目录下创建一个专用文件夹
    output_folder = os.path.join(os.path.dirname(file_path), "提取的图片")
    
    try:
        # 定义 Word XML 的命名空间
        ns = {
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'v': 'urn:schemas-microsoft-com:vml'
        }

        # docx 本质是 zip 压缩包，直接读取
        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            # 1. 解析关系文件，建立内部 ID (rId) 到真实图片路径的映射
            rels_xml = docx_zip.read('word/_rels/document.xml.rels')
            rels_tree = ET.fromstring(rels_xml)
            rels_map = {}
            for rel in rels_tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rels_map[rel.get('Id')] = rel.get('Target')

            # 2. 解析文档主体内容，这决定了图片从上到下的顺序
            doc_xml = docx_zip.read('word/document.xml')
            doc_tree = ET.fromstring(doc_xml)

            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            image_index = 1
            
            # 遍历文档节点树，寻找图片标签 (blip 或 imagedata)
            for elem in doc_tree.iter():
                r_id = None
                if elem.tag == f"{{{ns['a']}}}blip":
                    r_id = elem.get(f"{{{ns['r']}}}embed")
                elif elem.tag == f"{{{ns['v']}}}imagedata":
                    r_id = elem.get(f"{{{ns['r']}}}id")

                if r_id and r_id in rels_map:
                    target = rels_map[r_id]
                    # Target 通常是 'media/image1.jpeg' 这样的格式
                    target_clean = target.split('/')[-1]
                    zip_path = f"word/media/{target_clean}"
                    
                    try:
                        # 从压缩包内存中直接读取图片二进制数据
                        image_data = docx_zip.read(zip_path)
                        # 获取原始后缀名 (如 .jpg, .png)
                        ext = os.path.splitext(target_clean)[1]
                        if not ext:
                            ext = '.png'
                        
                        out_path = os.path.join(output_folder, f"图{image_index}{ext}")
                        with open(out_path, 'wb') as f:
                            f.write(image_data)
                        print(f"[成功] 提取 -> 图{image_index}{ext}")
                        image_index += 1
                    except KeyError:
                        pass # 忽略找不到的异常媒体节点
            
        print("="*50)
        if image_index > 1:
            print(f"提取完成！共成功提取 {image_index - 1} 张图片。")
            print(f"图片已保存在: {output_folder}")
        else:
            print("提取完成，但在文档中没有找到图片。")
            
    except zipfile.BadZipFile:
        print("\n[错误] 文件已损坏或不是有效的 .docx 文件。")
    except Exception as e:
        print(f"\n[程序异常] 发生未知错误: {e}")
        
    print("\n")
    os.system("pause") # 运行完毕后暂停，防止黑框闪退

if __name__ == '__main__':
    main()
