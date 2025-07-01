import pandas as pd
import os
import tempfile
from playwright.sync_api import sync_playwright
import re
from pypinyin import lazy_pinyin
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Entry
from tkinter.scrolledtext import ScrolledText
import threading
import queue
from concurrent.futures import ThreadPoolExecutor
import time


class FeedbackGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("课堂反馈表生成器")
        self.root.geometry("850x650")
        self.root.resizable(True, True)

        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure('.', font=('SimHei', 10))

        # 创建变量
        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.date_range = tk.StringVar()
        self.status = tk.StringVar()
        self.status.set("就绪")
        self.progress_var = tk.DoubleVar()

        # 列数变量
        self.col_student_name = tk.StringVar(value="4")  # 学生姓名列
        self.col_eng_name = tk.StringVar(value="6")  # 英文名列，默认第6列
        self.col_course_name = tk.StringVar(value="11")  # 课程名称列
        self.col_attendance = tk.StringVar(value="22")  # 出勤列
        self.col_performance = tk.StringVar(value="25")  # 课堂反馈列
        self.col_homework = tk.StringVar(value="26")  # 作业反馈列
        self.col_comment = tk.StringVar(value="27")  # 评语列

        # 学生英文名映射
        self.student_eng_names = {}

        # 创建界面
        self.create_widgets()

        # 默认值
        self.output_dir.set(os.path.join(os.getcwd(), "学生反馈PDF"))

        # 线程池和队列
        self.thread_pool = None
        self.task_queue = queue.Queue()
        self.running = False

    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Excel文件选择
        ttk.Label(main_frame, text="Excel文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file, width=50).grid(row=0, column=1, pady=5, padx=5)
        ttk.Button(main_frame, text="浏览...", command=self.browse_input_file).grid(row=0, column=2, pady=5)

        # 输出文件夹选择
        ttk.Label(main_frame, text="输出文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=1, column=1, pady=5, padx=5)
        ttk.Button(main_frame, text="浏览...", command=self.browse_output_dir).grid(row=1, column=2, pady=5)

        # 日期范围输入
        ttk.Label(main_frame, text="日期范围:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.date_range, width=50).grid(row=2, column=1, pady=5, padx=5)
        ttk.Label(main_frame, text="(例如：5月20日-5月24日)").grid(row=2, column=2, sticky=tk.W, pady=5)

        # 列数设置
        ttk.Label(main_frame, text="列数设置:").grid(row=3, column=0, sticky=tk.W, pady=5)

        ttk.Label(main_frame, text="学生姓名列:").grid(row=4, column=0, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.col_student_name, width=5).grid(row=4, column=1, pady=2, padx=5)
        ttk.Label(main_frame, text="英文名列:").grid(row=4, column=2, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.col_eng_name, width=5).grid(row=4, column=3, pady=2, padx=5)

        ttk.Label(main_frame, text="课程名称列:").grid(row=5, column=0, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.col_course_name, width=5).grid(row=5, column=1, pady=2, padx=5)
        ttk.Label(main_frame, text="出勤列:").grid(row=5, column=2, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.col_attendance, width=5).grid(row=5, column=3, pady=2, padx=5)

        ttk.Label(main_frame, text="课堂反馈列:").grid(row=6, column=0, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.col_performance, width=5).grid(row=6, column=1, pady=2, padx=5)
        ttk.Label(main_frame, text="作业反馈列:").grid(row=6, column=2, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.col_homework, width=5).grid(row=6, column=3, pady=2, padx=5)

        ttk.Label(main_frame, text="评语列:").grid(row=7, column=0, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=self.col_comment, width=5).grid(row=7, column=1, pady=2, padx=5)

        # 分隔线
        ttk.Separator(main_frame, orient="horizontal").grid(row=8, column=0, columnspan=4, sticky="ew", pady=15)

        # 状态标签
        ttk.Label(main_frame, text="状态:").grid(row=9, column=0, sticky=tk.W, pady=5)
        ttk.Label(main_frame, textvariable=self.status).grid(row=9, column=1, sticky=tk.W, pady=5)

        # 进度条
        ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100).grid(row=10, column=0, columnspan=4,
                                                                                  sticky="ew", pady=5)

        # 日志框
        ttk.Label(main_frame, text="处理日志:").grid(row=11, column=0, sticky=tk.W, pady=5)
        self.log_text = ScrolledText(main_frame, width=100, height=15)
        self.log_text.grid(row=12, column=0, columnspan=4, sticky="nsew", pady=5)

        # 按钮框架
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=13, column=0, columnspan=4, pady=15)

        self.start_btn = ttk.Button(btn_frame, text="开始生成", command=self.start_generation)
        self.start_btn.pack(side=tk.LEFT, padx=10)

        self.stop_btn = ttk.Button(btn_frame, text="停止", command=self.stop_generation, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=10)

        ttk.Button(btn_frame, text="退出", command=self.root.destroy).pack(side=tk.LEFT, padx=10)

        # 设置权重，使日志框可以拉伸
        main_frame.grid_rowconfigure(12, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_file.set(file_path)

    def browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择输出文件夹")
        if dir_path:
            self.output_dir.set(dir_path)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_generation(self):
        # 检查输入
        input_file = self.input_file.get()
        output_dir = self.output_dir.get()
        date_range = self.date_range.get()

        # 检查列数输入
        try:
            col_student = int(self.col_student_name.get())
            col_eng = int(self.col_eng_name.get())
            col_course = int(self.col_course_name.get())
            col_attendance = int(self.col_attendance.get())
            col_performance = int(self.col_performance.get())
            col_homework = int(self.col_homework.get())
            col_comment = int(self.col_comment.get())

            if not (0 <= col_student < 100 and 0 <= col_eng < 100 and
                    0 <= col_course < 100 and 0 <= col_attendance < 100 and
                    0 <= col_performance < 100 and 0 <= col_homework < 100 and
                    0 <= col_comment < 100):
                raise ValueError("列数必须在0-99之间")
        except ValueError as e:
            messagebox.showerror("错误", f"列数输入错误: {str(e)}")
            return

        if not input_file:
            messagebox.showerror("错误", "请选择Excel文件")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", f"文件不存在: {input_file}")
            return

        if not date_range:
            messagebox.showerror("错误", "请输入日期范围")
            return

        # 创建输出目录（如果不存在）
        os.makedirs(output_dir, exist_ok=True)

        # 清空日志和进度条
        self.log_text.delete(1.0, tk.END)
        self.progress_var.set(0)

        # 更新按钮状态
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        # 设置运行状态和列数
        self.running = True
        self.col_student = col_student
        self.col_eng = col_eng
        self.col_course = col_course
        self.col_attendance = col_attendance
        self.col_performance = col_performance
        self.col_homework = col_homework
        self.col_comment = col_comment

        # 清空英文名映射
        self.student_eng_names = {}

        # 在单独的线程中运行生成过程
        threading.Thread(target=self.generate_feedback, daemon=True).start()

    def stop_generation(self):
        self.running = False
        self.status.set("正在停止...")
        self.log("正在停止生成过程...")

        # 如果有线程池，关闭它
        if self.thread_pool:
            self.thread_pool.shutdown(wait=False)

        # 更新按钮状态
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def generate_feedback(self):
        try:
            self.status.set("正在处理...")
            input_file = self.input_file.get()
            output_dir = self.output_dir.get()
            date_range = self.date_range.get()

            self.log(f"开始处理Excel文件: {input_file}")
            self.log(f"输出文件夹: {output_dir}")
            self.log(f"日期范围: {date_range}")
            self.log(f"列数设置: 学生姓名={self.col_student}, 英文名={self.col_eng}, "
                     f"课程名称={self.col_course}, 出勤={self.col_attendance}, "
                     f"课堂反馈={self.col_performance}, 作业反馈={self.col_homework}, "
                     f"评语={self.col_comment}")

            # 处理Excel数据
            student_data = self.process_excel(input_file)
            total_students = len(student_data)
            self.log(f"共处理 {total_students} 名学生的数据")
            self.log(f"学生英文名映射已建立，共 {len(self.student_eng_names)} 条记录")

            if total_students == 0:
                self.log("没有可处理的学生数据")
                self.status.set("处理完成")
                self.start_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                return

            # 读取HTML模板
            base_dir = os.path.dirname(os.path.abspath(__file__))
            try:
                with open("structure.html", "r", encoding="utf-8") as f:
                    html_template = f.read()
            except FileNotFoundError:
                self.log("错误：找不到HTML模板文件 'structure.html'")
                self.status.set("处理失败")
                self.start_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                return

            # 创建线程池（根据CPU核心数设置线程数）
            max_workers = min(5, total_students)  # 最多5个线程
            self.log(f"使用 {max_workers} 个线程并行处理PDF")
            self.thread_pool = ThreadPoolExecutor(max_workers=max_workers)

            # 提交任务到线程池
            results = []
            for student_name, courses in student_data.items():
                if not self.running:
                    break

                if not courses:
                    self.log(f"跳过学生: {student_name}（无课程数据）")
                    continue

                future = self.thread_pool.submit(
                    self.generate_student_pdf,
                    html_template,
                    student_name,
                    courses,
                    date_range,
                    base_dir,
                    output_dir
                )
                results.append((student_name, future))

            # 等待所有任务完成并收集结果
            success_count = 0
            processed_count = 0

            for student_name, future in results:
                if not self.running:
                    break

                try:
                    success = future.result(timeout=300)  # 设置超时时间5分钟
                    if success:
                        success_count += 1
                    processed_count += 1
                    self.progress_var.set(100 * processed_count / total_students)
                except Exception as e:
                    self.log(f"处理 {student_name} 时出错: {str(e)}")

            # 关闭线程池
            self.thread_pool.shutdown()
            self.thread_pool = None

            # 更新状态
            if self.running:
                self.status.set(f"处理完成 ({success_count}/{total_students})")
                self.log(f"\n所有PDF已保存到 '{output_dir}' 文件夹")
                messagebox.showinfo("完成", f"处理完成！成功生成 {success_count} 个PDF文件")
            else:
                self.status.set(f"已停止 (处理 {processed_count}/{total_students})")
                self.log(f"\n生成过程已停止，部分PDF可能未生成")
                messagebox.showinfo("已停止", f"生成过程已停止，成功生成 {success_count} 个PDF文件")

        except Exception as e:
            self.status.set("处理失败")
            self.log(f"发生错误: {str(e)}")
            messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")
        finally:
            # 更新按钮状态
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

    def generate_student_pdf(self, html_template, student_name, courses, date_range, base_dir, output_dir):
        try:
            # 生成个性化HTML
            html_content = self.generate_dynamic_html(html_template, student_name, courses, date_range)
            html_content = self.make_image_paths_absolute(html_content, base_dir)

            # 生成PDF路径（修改此处：添加日期到文件名）
            # 处理日期格式，移除特殊字符并保留核心信息
            safe_date = re.sub(r'[\\/:*?"<>|]', '', date_range)  # 移除文件名非法字符
            pdf_name = f"{student_name}（{safe_date}）.pdf"
            pdf_path = os.path.join(output_dir, pdf_name)

            # 生成PDF
            self.generate_pdf_from_html(html_content, pdf_path)

            # 更新日志（使用线程安全的方式）
            self.root.after(0, lambda msg=f"成功生成: {pdf_path}": self.log(msg))

            return True
        except Exception as e:
            # 更新日志（使用线程安全的方式，显式绑定变量）
            error_msg = f"生成 {student_name} 的PDF时出错: {str(e)}"
            self.root.after(0, lambda msg=error_msg: self.log(msg))
            return False

    # 处理Excel文件，提取学生及其课程数据，并建立中文名与英文名的映射
    def process_excel(self, file_path):
        df = pd.read_excel(file_path)
        student_dict = {}

        for index, row in df.iterrows():
            # 过滤条件：排除艺术、音乐、咨询类课程
            if "ART" in str(row[0]) or "Music" in str(row[0]):
                continue
            if "Counseling" in str(row[self.col_course]) or "升学指导" in str(row[self.col_course]) :
                continue

            try:
                student_name = row[self.col_student]  # 学生姓名
                if pd.isna(student_name):
                    continue  # 跳过无姓名的行
            except:
                student_name = f"学生_{index}"
                self.log(f"警告：无法从第{self.col_student}列获取学生姓名，使用默认名称")

            # 获取英文名并建立映射
            try:
                eng_name = row[self.col_eng]
                if not pd.isna(eng_name):
                    self.student_eng_names[student_name] = str(eng_name).strip()
                else:
                    self.student_eng_names[student_name] = ""  # 无英文名时留空
            except:
                self.student_eng_names[student_name] = ""
                self.log(f"警告：无法从第{self.col_eng}列获取{student_name}的英文名")

            # 提取课程数据
            course_data = f"{row[self.col_course]},{row[self.col_attendance]},{row[self.col_performance]},{row[self.col_homework]},{row[self.col_comment]}"

            if student_name not in student_dict:
                student_dict[student_name] = []
            student_dict[student_name].append(course_data)

        return student_dict

    def name_to_pinyin(self, name):
        """将中文名转换为拼音（简拼）"""
        return ''.join(lazy_pinyin(name))

    def generate_dynamic_html(self, html_content, student_name, courses, date_range):

        # 获取学生英文名
        eng_name = self.student_eng_names.get(student_name, "")

        # 基础替换
        replace_rules = {
            "[name]": student_name,
            "[pinyin]": self.name_to_pinyin(student_name),
            "[engName]": eng_name,
            "[date]": date_range,
        }

        # 替换基础占位符
        for old_text, new_text in replace_rules.items():
            html_content = html_content.replace(old_text, new_text)

        # 动态生成课程部分HTML
        course_html = ""
        for i, course_str in enumerate(courses, 1):
            # 分割课程数据字符串
            parts = course_str.split(',')

            # 确保数据完整性（至少需要5个字段，包含评语）
            if len(parts) >= 5:
                course_name = parts[0]
                try:
                    attendance = int(float(parts[1]))  # 先转浮点数，再取整
                    performance = int(float(parts[2]))
                    homework = int(float(parts[3]))
                except (ValueError, TypeError):  # 处理非数字或空值
                    attendance = 0
                    performance = 0
                    homework = 0
                comment = parts[4]
            else:
                # 数据不足时的默认值
                course_name = f"课程{i}"
                attendance = "N/A"
                performance = "N/A"
                homework = "N/A"
                comment = "待评价"

            # 生成单个课程的HTML（关键修改点：评语独立分行）
            course_html += f"""
    <div class="course-block">
        <div class="course-header">
            <div class="course-name">
                <p>{course_name}</p>
            </div>
            <div class="course-ratings">
                <div class="rating-item">
                    <p>出勤：<span class="content">{attendance}</span></p>
                </div>
                <div class="rating-item">
                    <p>课堂反馈：<span class="content">{performance}</span></p>
                </div>
                <div class="rating-item">
                    <p>作业反馈：<span class="content">{homework}</span></p>
                </div>
            </div>
        </div>
        <div class="course-comment">
            <p>评语：<span class="content">{comment}</span></p>
        </div>
    </div>
    <div class="dashed-separator"></div>
    """

        # 替换课程占位符
        html_content = html_content.replace("<div id=\"course-section\"></div>",
                                            f"<div id=\"course-section\">{course_html}</div>")
        return html_content

    def make_image_paths_absolute(self, html_content, base_dir):
        """转换图片路径为绝对路径"""
        img_pattern = re.compile(r'<img\s+[^>]*src="([^"]+)"[^>]*>')
        return img_pattern.sub(
            lambda m: f'<img src="{os.path.join(base_dir, m.group(1))}" alt="爱迪学校logo" class="logo">',
            html_content
        )

    def generate_pdf_from_html(self, html_string, pdf_path):
        """生成PDF文件"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
            temp_path = f.name
            f.write(html_string)

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=True,
                    args=["--allow-file-access-from-files"]
                )
                page = browser.new_page()
                file_url = f"file://{temp_path}"
                page.goto(file_url, wait_until="networkidle")

                # 等待图片加载
                page.wait_for_selector("img", state="visible", timeout=5000)

                # 滚动页面确保内容完整
                page_height = page.evaluate('document.body.scrollHeight')
                scroll_step = 500
                total_steps = max(1, page_height // scroll_step)  # 至少滚动一次
                for step in range(total_steps):
                    scroll_position = step * scroll_step
                    page.evaluate(f'window.scrollTo(0, {scroll_position});')
                    page.wait_for_timeout(100)

                # 生成PDF
                page.pdf(
                    path=pdf_path,
                    format="A4",
                    print_background=True,
                    prefer_css_page_size=True,
                    scale=1.0
                )
        finally:
            os.unlink(temp_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = FeedbackGeneratorApp(root)
    root.mainloop()
