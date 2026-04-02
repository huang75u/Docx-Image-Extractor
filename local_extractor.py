import zipfile
import xml.etree.ElementTree as ET
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def log_message(text_widget, message):
    """向界面日志窗口输出信息"""
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END) # 自动滚动到最底部
    text_widget.update()    # 强制刷新界面

def extract_images(file_path, text_widget):
    if not file_path.lower().endswith('.docx'):
        log_message(text_widget, "[错误] 请确保提供的是 .docx 格式的文件！")
        return

    output_folder = os.path.join(os.path.dirname(file_path), "提取的图片")
    log_message(text_widget, f"开始解析文档: {os.path.basename(file_path)}")
    
    try:
        ns = {
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'v': 'urn:schemas-microsoft-com:vml'
        }

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            rels_xml = docx_zip.read('word/_rels/document.xml.rels')
            rels_tree = ET.fromstring(rels_xml)
            rels_map = {}
            for rel in rels_tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rels_map[rel.get('Id')] = rel.get('Target')

            doc_xml = docx_zip.read('word/document.xml')
            doc_tree = ET.fromstring(doc_xml)

            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            image_index = 1
            
            for elem in doc_tree.iter():
                r_id = None
                if elem.tag == f"{{{ns['a']}}}blip":
                    r_id = elem.get(f"{{{ns['r']}}}embed")
                elif elem.tag == f"{{{ns['v']}}}imagedata":
                    r_id = elem.get(f"{{{ns['r']}}}id")

                if r_id and r_id in rels_map:
                    target = rels_map[r_id]
                    target_clean = target.split('/')[-1]
                    zip_path = f"word/media/{target_clean}"
                    
                    try:
                        image_data = docx_zip.read(zip_path)
                        ext = os.path.splitext(target_clean)[1]
                        if not ext:
                            ext = '.png'
                        
                        out_path = os.path.join(output_folder, f"图{image_index}{ext}")
                        with open(out_path, 'wb') as f:
                            f.write(image_data)
                        log_message(text_widget, f"  [成功] 提取 -> 图{image_index}{ext}")
                        image_index += 1
                    except KeyError:
                        pass
            
        if image_index > 1:
            log_message(text_widget, "-"*40)
            log_message(text_widget, f"🎉 提取完成！共成功提取 {image_index - 1} 张图片。")
            log_message(text_widget, f"📂 图片已保存在: {output_folder}")
            messagebox.showinfo("完成", f"成功提取 {image_index - 1} 张图片！\n\n文件夹已生成在文档所在目录。")
        else:
            log_message(text_widget, "提取完成，但在文档中没有找到图片。")
            messagebox.showwarning("提示", "文档中没有找到图片。")
            
    except zipfile.BadZipFile:
        log_message(text_widget, "[错误] 文件已损坏或不是有效的 .docx 文件。")
        messagebox.showerror("错误", "文件已损坏或不是有效的 .docx 文件。")
    except Exception as e:
        log_message(text_widget, f"[程序异常] 发生未知错误: {e}")
        messagebox.showerror("错误", f"发生未知错误: {e}")

def select_file_and_run(text_widget):
    # 弹出文件选择框，并设置默认打开的路径
    file_path = filedialog.askopenfilename(
        title="选择一个 Word (.docx) 文件",
        initialdir=r"D:\Huang_Work_Space\自用工具\存图的doc", # <--- 这里添加了默认路径
        filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
    )
    if file_path:
        # 清空上次的日志
        text_widget.delete(1.0, tk.END)
        extract_images(file_path, text_widget)

def main():
    # 创建主窗口
    root = tk.Tk()
    root.title("Word 图片顺序提取工具")
    root.geometry("500x380")
    root.resizable(False, False) # 禁止拉伸窗口

    # 标题标签
    title_label = tk.Label(root, text="📄 Word (.docx) 图片提取器", font=("微软雅黑", 16, "bold"))
    title_label.pack(pady=15)

    # 提示说明
    desc_label = tk.Label(root, text="选择一个 .docx 文件，程序将按文档从上到下的顺序提取图片", font=("微软雅黑", 10), fg="gray")
    desc_label.pack(pady=5)

    # 日志文本框 (带滚动条)
    frame = tk.Frame(root)
    frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
    
    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    log_text = tk.Text(frame, height=10, yscrollcommand=scrollbar.set, font=("Consolas", 9), bg="#f4f4f4")
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=log_text.yview)
    
    log_message(log_text, "等待选择文件...")

    # 选择文件并提取的按钮
    btn = tk.Button(root, text="选择 .docx 文件并开始提取", font=("微软雅黑", 11, "bold"), bg="#4CAF50", fg="white", 
                    activebackground="#45a049", activeforeground="white", cursor="hand2",
                    command=lambda: select_file_and_run(log_text))
    btn.pack(pady=15, ipadx=10, ipady=5)

    # 启动界面循环
    root.mainloop()

if __name__ == '__main__':
    main()
