import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, scrolledtext
import os
from pathlib import Path
import subprocess
import platform
import shutil

# 导入PDF处理库
try:
    from PyPDF2 import PdfWriter, PdfReader, Transformation, PageObject
except ImportError as e:
    print(f"导入PyPDF2库失败: {e}")
    print("请确保已正确安装PyPDF2库: pip install PyPDF2")
    exit(1)

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
except ImportError as e:
    print(f"导入reportlab库失败: {e}")
    print("请确保已正确安装reportlab库: pip install reportlab")
    exit(1)


class PDFResizerModernGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF页面缩放工具 - 现代版")
        self.root.geometry("700x600")
        self.root.minsize(600, 500)
        
        # 设置界面变量
        self.input_folder = ttk.StringVar()
        self.output_folder = ttk.StringVar()
        self.max_width = ttk.DoubleVar(value=10.0)
        self.max_height = ttk.DoubleVar(value=15.0)
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=BOTH, expand=YES)
        
        # 标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=X, pady=(0, 20))
        
        title_label = ttk.Label(
            title_frame, 
            text="PDF页面缩放工具", 
            font=("Arial", 20, "bold"),
            bootstyle="primary"
        )
        title_label.pack(side=LEFT)
        
        # 说明卡片
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=X, pady=(0, 20))
        
        info_text = "将PDF页面按比例缩放，使页面尺寸适应指定范围\n" \
                   "缩放规则:\n" \
                   "- 页面按比例缩放\n" \
                   "- 宽度最大不超过设定值，高度最大不超过设定值\n" \
                   "- 例如：100*150cm -> 10*15cm, 100*1500cm -> 1*15cm, 1000*150cm -> 10*1.5cm"
        
        # 使用Frame替代Card，并添加样式
        info_card = ttk.Frame(info_frame, relief=RIDGE, borderwidth=2)
        info_card.pack(fill=X, padx=5, pady=5)
        
        info_label = ttk.Label(
            info_card, 
            text=info_text, 
            bootstyle="secondary",
            justify=LEFT,
            padding=10
        )
        info_label.pack(fill=X)
        
        # 文件夹选择区域
        folder_frame = ttk.Frame(main_frame)
        folder_frame.pack(fill=X, pady=(0, 20))
        
        # 输入文件夹
        input_group = ttk.Labelframe(folder_frame, text="输入设置", padding=10)
        input_group.pack(fill=X, pady=(0, 10))
        
        input_row = ttk.Frame(input_group)
        input_row.pack(fill=X, pady=5)
        
        ttk.Label(input_row, text="输入文件夹:", width=12, anchor=W).pack(side=LEFT)
        
        self.input_entry = ttk.Entry(input_row, textvariable=self.input_folder)
        self.input_entry.pack(side=LEFT, fill=X, expand=YES, padx=5)
        
        ttk.Button(
            input_row, 
            text="浏览", 
            command=self.browse_input_folder,
            bootstyle="info"
        ).pack(side=LEFT)
        
        # 输出文件夹
        output_row = ttk.Frame(input_group)
        output_row.pack(fill=X, pady=5)
        
        ttk.Label(output_row, text="输出文件夹:", width=12, anchor=W).pack(side=LEFT)
        
        self.output_entry = ttk.Entry(output_row, textvariable=self.output_folder)
        self.output_entry.pack(side=LEFT, fill=X, expand=YES, padx=5)
        
        ttk.Button(
            output_row, 
            text="浏览", 
            command=self.browse_output_folder,
            bootstyle="info"
        ).pack(side=LEFT)
        
        # 尺寸设置区域
        size_frame = ttk.Labelframe(main_frame, text="尺寸设置 (厘米)", padding=10)
        size_frame.pack(fill=X, pady=(0, 20))
        
        size_row = ttk.Frame(size_frame)
        size_row.pack(fill=X)
        
        ttk.Label(size_row, text="最大宽度:", width=12, anchor=W).pack(side=LEFT)
        
        width_spinbox = ttk.Spinbox(
            size_row, 
            from_=1, 
            to=50, 
            increment=0.1, 
            textvariable=self.max_width, 
            width=10
        )
        width_spinbox.pack(side=LEFT, padx=5)
        
        ttk.Label(size_row, text="cm", width=5, anchor=W).pack(side=LEFT)
        
        ttk.Label(size_row, text="最大高度:", width=12, anchor=E).pack(side=LEFT, padx=(20, 0))
        
        height_spinbox = ttk.Spinbox(
            size_row, 
            from_=1, 
            to=50, 
            increment=0.1, 
            textvariable=self.max_height, 
            width=10
        )
        height_spinbox.pack(side=LEFT, padx=5)
        
        ttk.Label(size_row, text="cm", width=5, anchor=W).pack(side=LEFT)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(0, 20))
        
        self.create_sample_btn = ttk.Button(
            button_frame,
            text="创建示例文件",
            command=self.create_sample,
            bootstyle="success"
        )
        self.create_sample_btn.pack(side=LEFT, padx=(0, 10))
        
        self.process_btn = ttk.Button(
            button_frame,
            text="开始处理",
            command=self.process_pdfs,
            bootstyle="primary"
        )
        self.process_btn.pack(side=LEFT, padx=(0, 10))
        
        self.open_folder_btn = ttk.Button(
            button_frame,
            text="打开输出文件夹",
            command=self.open_output_folder,
            bootstyle="info",
            state=DISABLED
        )
        self.open_folder_btn.pack(side=LEFT)
        
        # 日志区域
        log_frame = ttk.Labelframe(main_frame, text="处理日志", padding=5)
        log_frame.pack(fill=BOTH, expand=YES)
        
        log_frame_inner = ttk.Frame(log_frame)
        log_frame_inner.pack(fill=BOTH, expand=YES)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame_inner, 
            height=12,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=BOTH, expand=YES)
        
        # 状态栏
        self.status_var = ttk.StringVar(value="就绪")
        status_bar = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            bootstyle="secondary",
            padding=(0, 5)
        )
        status_bar.pack(fill=X)
        
    def open_output_folder(self):
        """打开输出文件夹"""
        output_folder = self.output_folder.get().strip()
        if output_folder and os.path.exists(output_folder):
            # 根据操作系统打开文件夹
            try:
                if platform.system() == "Windows":
                    os.startfile(output_folder)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.Popen(["open", output_folder])
                else:  # Linux
                    subprocess.Popen(["xdg-open", output_folder])
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")
        else:
            messagebox.showwarning("警告", "输出文件夹不存在或未设置")
        
    def browse_input_folder(self):
        folder = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if folder:
            self.input_folder.set(folder)
            self.status_var.set(f"已选择输入文件夹: {folder}")
            
    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_folder.set(folder)
            self.status_var.set(f"已选择输出文件夹: {folder}")
            
    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()
        
    def resize_pdf_pages(self, input_path, output_path, max_width_cm=10, max_height_cm=15):
        """
        将PDF的每一页按比例缩放到指定尺寸范围内
        """
        try:
            # 读取输入PDF
            reader = PdfReader(input_path)
            writer = PdfWriter()
            
            # 转换厘米到点(point)
            # 1cm = 28.3464567 points
            max_width_points = max_width_cm * 28.3464567
            max_height_points = max_height_cm * 28.3464567
            
            for page_num, page in enumerate(reader.pages):
                # 获取原始页面尺寸（单位：点point）
                original_width = float(page.mediabox.width)
                original_height = float(page.mediabox.height)
                
                # 计算宽度和高度的缩放因子
                width_scale = max_width_points / original_width
                height_scale = max_height_points / original_height
                
                # 选择较小的缩放因子以保持比例并适应两个维度
                scale_factor = min(width_scale, height_scale)
                
                # 计算新尺寸
                new_width = original_width * scale_factor
                new_height = original_height * scale_factor
                
                # 创建具有新尺寸的空白页面
                new_page = PageObject.create_blank_page(
                    width=new_width,
                    height=new_height
                )
                
                # 使用现代API进行缩放：先缩放原页面，再合并
                page.scale_by(scale_factor)
                new_page.merge_page(page)
                
                # 将新页面添加到writer
                writer.add_page(new_page)
            
            # 写入输出文件
            with open(output_path, "wb") as output_file:
                writer.write(output_file)
                
            return True
        except Exception as e:
            self.log(f"  缩放失败: {str(e)}")
            return False
            
    def process_multiple_pdfs(self, input_folder, output_folder, max_width_cm=10, max_height_cm=15):
        """
        批量处理文件夹中的PDF文件
        """
        try:
            # 确保输出文件夹存在
            Path(output_folder).mkdir(parents=True, exist_ok=True)
            
            # 遍历输入文件夹中的所有PDF文件
            input_path = Path(input_folder)
            pdf_files = list(input_path.glob("*.pdf"))
            
            if not pdf_files:
                self.log(f"在 {input_folder} 中未找到PDF文件")
                return False
                
            self.log(f"找到 {len(pdf_files)} 个PDF文件")
            
            success_count = 0
            for pdf_file in pdf_files:
                output_file = Path(output_folder) / pdf_file.name  # 使用原文件名
                self.log(f"正在处理: {pdf_file.name}")
                try:
                    if self.resize_pdf_pages(str(pdf_file), str(output_file), max_width_cm, max_height_cm):
                        self.log(f"  成功: {pdf_file.name}")
                        success_count += 1
                    else:
                        self.log(f"  失败: {pdf_file.name}")
                except Exception as e:
                    self.log(f"  处理 {pdf_file.name} 时出错: {str(e)}")
                    
            self.log(f"处理完成! 成功处理 {success_count}/{len(pdf_files)} 个文件")
            return True
            
        except Exception as e:
            self.log(f"批量处理过程中发生错误: {str(e)}")
            return False
            
    def create_sample_pdf(self, filename):
        """
        创建示例PDF用于测试
        """
        try:
            c = canvas.Canvas(filename)
            c.setFont("Helvetica", 12)
            
            # 创建一个较大的页面 (模拟各种尺寸)
            c.setPageSize((300, 400))  # 约10.59cm x 14.17cm
            c.drawString(50, 350, "这是一个测试PDF文件")
            c.drawString(50, 330, "原始尺寸: 约10.59cm x 14.17cm")
            c.drawString(50, 310, "将会被缩放到指定尺寸范围内")
            c.rect(50, 100, 200, 200)  # 绘制一个矩形
            c.drawString(60, 200, "这是页面上的一个矩形")
            c.showPage()
            
            # 第二页
            c.setPageSize((600, 800))  # 约21.18cm x 28.35cm (超大页面)
            c.drawString(50, 750, "第二页 - 更大的页面")
            c.drawString(50, 730, "原始尺寸: 约21.18cm x 28.35cm")
            c.drawString(50, 710, "将会按比例缩小")
            c.rect(100, 200, 400, 400)  # 较大的矩形
            c.drawString(120, 400, "这是第二页的大矩形")
            c.showPage()
            
            c.save()
            return True
        except Exception as e:
            self.log(f"创建示例文件时出错: {str(e)}")
            return False
            
    def create_sample(self):
        """
        创建示例文件
        """
        try:
            filename = "sample_input.pdf"
            if self.create_sample_pdf(filename):
                self.log(f"示例PDF文件 '{filename}' 已创建")
                messagebox.showinfo("成功", f"示例文件 {filename} 已创建!")
                self.status_var.set(f"已创建示例文件: {filename}")
            else:
                messagebox.showerror("错误", "创建示例文件失败")
        except Exception as e:
            messagebox.showerror("错误", f"创建示例文件时出错: {str(e)}")
            
    def process_pdfs(self):
        """
        开始处理PDF文件
        """
        # 检查输入参数
        input_folder = self.input_folder.get().strip()
        output_folder = self.output_folder.get().strip()
        
        if not input_folder:
            messagebox.showwarning("警告", "请选择输入文件夹")
            return
            
        if not output_folder:
            messagebox.showwarning("警告", "请选择输出文件夹")
            return
            
        if not os.path.exists(input_folder):
            messagebox.showerror("错误", "输入文件夹不存在")
            return
            
        if not os.path.isdir(input_folder):
            messagebox.showerror("错误", "输入路径不是一个文件夹")
            return
            
        # 获取尺寸参数
        max_width_cm = self.max_width.get()
        max_height_cm = self.max_height.get()
        
        # 开始处理
        self.process_btn.config(state=DISABLED)
        self.create_sample_btn.config(state=DISABLED)
        self.open_folder_btn.config(state=DISABLED)
        self.status_var.set("正在处理PDF文件...")
        
        self.log("开始处理PDF文件...")
        self.log(f"最大宽度: {max_width_cm}cm, 最大高度: {max_height_cm}cm")
        
        try:
            if self.process_multiple_pdfs(input_folder, output_folder, max_width_cm, max_height_cm):
                self.log("所有文件处理完成!")
                messagebox.showinfo("完成", "PDF文件处理完成!")
                self.status_var.set("处理完成")
                # 处理成功后启用打开文件夹按钮
                self.open_folder_btn.config(state=NORMAL)
            else:
                messagebox.showerror("错误", "处理过程中发生错误")
                self.status_var.set("处理出错")
        except Exception as e:
            self.log(f"处理过程中发生异常: {str(e)}")
            messagebox.showerror("错误", f"处理过程中发生异常: {str(e)}")
            self.status_var.set("处理异常")
        finally:
            self.process_btn.config(state=NORMAL)
            self.create_sample_btn.config(state=NORMAL)


def main():
    # 创建 ttkbootstrap 窗口，使用现代主题
    root = ttk.Window(themename="cosmo")  # 可选主题: cosmo, flatly, litera, minty, lumen, sandstone, yeti, pulse
    app = PDFResizerModernGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()