"""
Description: This script automates video course playback by recording mouse actions with interval settings.
It provides a GUI for users to set start and play positions, intervals, and number of courses
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pyautogui
import time
import json
import threading
from datetime import datetime

class VideoCourseAutomator:
    def __init__(self, root):
        self.root = root
        self.root.title("视频课程自动化工具 - 间隔记录版")
        self.root.geometry("950x700")
        
        # 脚本步骤存储
        self.script_steps = []
        self.is_recording = False
        self.is_playing = False
        self.current_step_index = 0
        
        self.setup_ui()
        self.update_mouse_position()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 鼠标位置显示
        pos_frame = ttk.LabelFrame(main_frame, text="鼠标位置监控", padding="5")
        pos_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.position_label = ttk.Label(pos_frame, text="X: 0, Y: 0", font=("Arial", 12))
        self.position_label.pack(pady=5)
        
        # 间隔记录区域
        interval_frame = ttk.LabelFrame(main_frame, text="间隔记录设置", padding="5")
        interval_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 第一行：起始位置和播放位置
        ttk.Label(interval_frame, text="起始X:").grid(row=0, column=0, padx=2)
        self.start_x = ttk.Entry(interval_frame, width=8)
        self.start_x.grid(row=0, column=1, padx=2)
        
        ttk.Label(interval_frame, text="起始Y:").grid(row=0, column=2, padx=2)
        self.start_y = ttk.Entry(interval_frame, width=8)
        self.start_y.grid(row=0, column=3, padx=2)
        
        ttk.Label(interval_frame, text="播放X:").grid(row=0, column=4, padx=2)
        self.play_x = ttk.Entry(interval_frame, width=8)
        self.play_x.grid(row=0, column=5, padx=2)
        
        ttk.Label(interval_frame, text="播放Y:").grid(row=0, column=6, padx=2)
        self.play_y = ttk.Entry(interval_frame, width=8)
        self.play_y.grid(row=0, column=7, padx=2)
        
        # 第二行：间隔设置
        ttk.Label(interval_frame, text="X间隔:").grid(row=1, column=0, padx=2)
        self.interval_x = ttk.Entry(interval_frame, width=8)
        self.interval_x.grid(row=1, column=1, padx=2)
        
        ttk.Label(interval_frame, text="Y间隔:").grid(row=1, column=2, padx=2)
        self.interval_y = ttk.Entry(interval_frame, width=8)
        self.interval_y.grid(row=1, column=3, padx=2)
        
        ttk.Label(interval_frame, text="课程数量:").grid(row=1, column=4, padx=2)
        self.course_count = ttk.Spinbox(interval_frame, from_=1, to=100, width=8)
        self.course_count.grid(row=1, column=5, padx=2)
        self.course_count.set("5")
        
        ttk.Label(interval_frame, text="视频时长(秒):").grid(row=1, column=6, padx=2)
        self.video_duration = ttk.Entry(interval_frame, width=10)
        self.video_duration.grid(row=1, column=7, padx=2)
        self.video_duration.insert(0, "300")
        
        # 第三行：按钮
        ttk.Button(interval_frame, text="获取起始位置", 
                  command=self.get_start_position).grid(row=2, column=0, columnspan=2, padx=2, pady=5)
        ttk.Button(interval_frame, text="获取播放位置", 
                  command=self.get_play_position).grid(row=2, column=2, columnspan=2, padx=2, pady=5)
        ttk.Button(interval_frame, text="生成间隔步骤", 
                  command=self.generate_interval_steps).grid(row=2, column=4, columnspan=2, padx=2, pady=5)
        ttk.Button(interval_frame, text="预览步骤", 
                  command=self.preview_interval_positions).grid(row=2, column=6, columnspan=2, padx=2, pady=5)
        
        # 脚本控制区域
        control_frame = ttk.LabelFrame(main_frame, text="脚本控制", padding="5")
        control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 回放按钮
        self.play_btn = ttk.Button(control_frame, text="执行脚本", command=self.execute_script)
        self.play_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # 停止按钮
        self.stop_btn = ttk.Button(control_frame, text="停止", command=self.stop_script)
        self.stop_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # 清除按钮
        clear_btn = ttk.Button(control_frame, text="清除脚本", command=self.clear_script)
        clear_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 测试按钮
        test_btn = ttk.Button(control_frame, text="测试步骤", command=self.test_current_step)
        test_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # 脚本步骤编辑区域
        script_frame = ttk.LabelFrame(main_frame, text="脚本步骤编辑", padding="5")
        script_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 步骤列表
        self.steps_tree = ttk.Treeview(script_frame, columns=("序号", "类型", "X", "Y", "等待时间", "描述"), show="headings", height=10)
        self.steps_tree.heading("序号", text="序号")
        self.steps_tree.heading("类型", text="类型")
        self.steps_tree.heading("X", text="X坐标")
        self.steps_tree.heading("Y", text="Y坐标")
        self.steps_tree.heading("等待时间", text="等待时间(秒)")
        self.steps_tree.heading("描述", text="描述")
        
        self.steps_tree.column("序号", width=50)
        self.steps_tree.column("类型", width=80)
        self.steps_tree.column("X", width=80)
        self.steps_tree.column("Y", width=80)
        self.steps_tree.column("等待时间", width=100)
        self.steps_tree.column("描述", width=250)
        
        self.steps_tree.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(script_frame, orient=tk.VERTICAL, command=self.steps_tree.yview)
        scrollbar.grid(row=0, column=4, sticky=(tk.N, tk.S))
        self.steps_tree.configure(yscrollcommand=scrollbar.set)
        
        # 步骤编辑区域
        edit_frame = ttk.Frame(script_frame)
        edit_frame.grid(row=1, column=0, columnspan=5, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(edit_frame, text="步骤类型:").grid(row=0, column=0, padx=2)
        self.step_type = ttk.Combobox(edit_frame, values=["点击", "等待", "移动"], state="readonly")
        self.step_type.grid(row=0, column=1, padx=2)
        self.step_type.set("点击")
        
        ttk.Label(edit_frame, text="X坐标:").grid(row=0, column=2, padx=2)
        self.x_entry = ttk.Entry(edit_frame, width=8)
        self.x_entry.grid(row=0, column=3, padx=2)
        
        ttk.Label(edit_frame, text="Y坐标:").grid(row=0, column=4, padx=2)
        self.y_entry = ttk.Entry(edit_frame, width=8)
        self.y_entry.grid(row=0, column=5, padx=2)
        
        ttk.Label(edit_frame, text="等待时间(秒):").grid(row=0, column=6, padx=2)
        self.wait_entry = ttk.Entry(edit_frame, width=8)
        self.wait_entry.grid(row=0, column=7, padx=2)
        self.wait_entry.insert(0, "5")
        
        ttk.Label(edit_frame, text="描述:").grid(row=0, column=8, padx=2)
        self.desc_entry = ttk.Entry(edit_frame, width=20)
        self.desc_entry.grid(row=0, column=9, padx=2)
        
        # 获取当前位置按钮
        get_pos_btn = ttk.Button(edit_frame, text="获取当前位置", command=self.get_current_position)
        get_pos_btn.grid(row=0, column=10, padx=5)
        
        # 添加步骤按钮
        add_step_btn = ttk.Button(edit_frame, text="添加步骤", command=self.add_step)
        add_step_btn.grid(row=0, column=11, padx=5)
        
        # 删除步骤按钮
        del_step_btn = ttk.Button(edit_frame, text="删除选中步骤", command=self.delete_step)
        del_step_btn.grid(row=0, column=12, padx=5)
        
        # 文件操作区域
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(file_frame, text="保存脚本", command=self.save_script).grid(row=0, column=0, padx=5)
        ttk.Button(file_frame, text="加载脚本", command=self.load_script).grid(row=0, column=1, padx=5)
        
        # 执行设置区域
        settings_frame = ttk.LabelFrame(main_frame, text="执行设置", padding="5")
        settings_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(settings_frame, text="循环次数:").grid(row=0, column=0, padx=2)
        self.loop_count = ttk.Spinbox(settings_frame, from_=1, to=1000, width=8)
        self.loop_count.grid(row=0, column=1, padx=2)
        self.loop_count.set("1")
        
        ttk.Label(settings_frame, text="循环间隔(秒):").grid(row=0, column=2, padx=2)
        self.loop_interval = ttk.Entry(settings_frame, width=8)
        self.loop_interval.grid(row=0, column=3, padx=2)
        self.loop_interval.insert(0, "2")
        
        # 状态显示
        self.status_label = ttk.Label(main_frame, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 配置权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        script_frame.columnconfigure(0, weight=1)
        script_frame.rowconfigure(0, weight=1)
        
    def update_mouse_position(self):
        """实时更新鼠标位置"""
        if not self.is_playing:  # 只在非播放状态更新位置
            x, y = pyautogui.position()
            self.position_label.config(text=f"X: {x}, Y: {y}")
        self.root.after(100, self.update_mouse_position)
    
    def get_current_position(self):
        """获取当前鼠标位置并填入输入框"""
        x, y = pyautogui.position()
        self.x_entry.delete(0, tk.END)
        self.x_entry.insert(0, str(x))
        self.y_entry.delete(0, tk.END)
        self.y_entry.insert(0, str(y))
    
    def get_start_position(self):
        """获取起始位置"""
        x, y = pyautogui.position()
        self.start_x.delete(0, tk.END)
        self.start_x.insert(0, str(x))
        self.start_y.delete(0, tk.END)
        self.start_y.insert(0, str(y))
        self.set_status(f"已设置起始位置: ({x}, {y})")
    
    def get_play_position(self):
        """获取播放位置"""
        x, y = pyautogui.position()
        self.play_x.delete(0, tk.END)
        self.play_x.insert(0, str(x))
        self.play_y.delete(0, tk.END)
        self.play_y.insert(0, str(y))
        self.set_status(f"已设置播放位置: ({x}, {y})")
    
    def generate_interval_steps(self):
        """生成间隔步骤：奇数序号选择视频，偶数序号播放"""
        try:
            start_x = int(self.start_x.get())
            start_y = int(self.start_y.get())
            play_x = int(self.play_x.get())
            play_y = int(self.play_y.get())
            interval_x = int(self.interval_x.get()) if self.interval_x.get() else 0
            interval_y = int(self.interval_y.get()) if self.interval_y.get() else 0
            course_count = int(self.course_count.get())
            video_duration = float(self.video_duration.get())
            
            if course_count <= 0:
                messagebox.showerror("错误", "课程数量必须大于0")
                return
            
            # 清除现有步骤
            if messagebox.askyesno("确认", "是否清除现有步骤并生成新的间隔步骤？"):
                self.script_steps.clear()
            
            # 生成间隔步骤
            for i in range(course_count):
                # 计算当前课程的位置
                course_x = start_x + (interval_x * i)
                course_y = start_y + (interval_y * i)
                
                # 奇数序号：选择视频
                select_step = {
                    "type": "点击",
                    "x": course_x,
                    "y": course_y,
                    "wait": 2,  # 选择后等待2秒加载
                    "desc": f"选择视频{i+1}"
                }
                self.script_steps.append(select_step)
                
                # 偶数序号：播放视频
                play_step = {
                    "type": "点击",
                    "x": play_x,
                    "y": play_y,
                    "wait": video_duration,  # 播放等待时间
                    "desc": f"播放视频{i+1}"
                }
                self.script_steps.append(play_step)
            
            self.update_steps_display()
            total_steps = len(self.script_steps)
            self.set_status(f"已生成 {total_steps} 个步骤 ({course_count}个课程)")
            
        except ValueError as e:
            messagebox.showerror("错误", "请输入有效的数字参数")
    
    def preview_interval_positions(self):
        """预览间隔位置和步骤"""
        try:
            start_x = int(self.start_x.get())
            start_y = int(self.start_y.get())
            play_x = int(self.play_x.get())
            play_y = int(self.play_y.get())
            interval_x = int(self.interval_x.get()) if self.interval_x.get() else 0
            interval_y = int(self.interval_y.get()) if self.interval_y.get() else 0
            course_count = int(self.course_count.get())
            video_duration = float(self.video_duration.get())
            
            preview_window = tk.Toplevel(self.root)
            preview_window.title("步骤预览")
            preview_window.geometry("500x600")
            
            text_widget = scrolledtext.ScrolledText(preview_window, width=60, height=35)
            text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            
            preview_text = f"步骤预览 (共{course_count}个课程):\n{'='*50}\n\n"
            preview_text += f"播放按钮位置: ({play_x}, {play_y})\n"
            preview_text += f"视频时长: {video_duration}秒\n\n"
            
            step_number = 1
            for i in range(course_count):
                course_x = start_x + (interval_x * i)
                course_y = start_y + (interval_y * i)
                
                # 选择步骤
                preview_text += f"步骤{step_number}: 选择视频{i+1}\n"
                preview_text += f"   位置: ({course_x}, {course_y})\n"
                preview_text += f"   等待: 2秒\n\n"
                step_number += 1
                
                # 播放步骤
                preview_text += f"步骤{step_number}: 播放视频{i+1}\n"
                preview_text += f"   位置: ({play_x}, {play_y})\n"
                preview_text += f"   等待: {video_duration}秒\n\n"
                step_number += 1
            
            total_time = (video_duration + 2) * course_count
            preview_text += f"{'='*50}\n"
            preview_text += f"总步骤数: {step_number-1}\n"
            preview_text += f"预计总时间: {total_time/60:.1f}分钟\n"
            
            text_widget.insert(tk.END, preview_text)
            text_widget.config(state=tk.DISABLED)
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字参数")
    
    def test_current_step(self):
        """测试当前选中的步骤"""
        selection = self.steps_tree.selection()
        if selection:
            index = int(selection[0]) - 1
            if 0 <= index < len(self.script_steps):
                step = self.script_steps[index]
                self.set_status(f"测试步骤: {step['desc']}")
                
                if step["type"] == "点击":
                    pyautogui.click(step["x"], step["y"])
                    self.set_status(f"已测试点击: ({step['x']}, {step['y']})")
                elif step["type"] == "移动":
                    pyautogui.moveTo(step["x"], step["y"])
                    self.set_status(f"已测试移动: ({step['x']}, {step['y']})")
        else:
            messagebox.showwarning("警告", "请先选择一个步骤进行测试")
    
    def add_step(self):
        """添加步骤到脚本"""
        try:
            step_type = self.step_type.get()
            x = int(self.x_entry.get()) if self.x_entry.get() else 0
            y = int(self.y_entry.get()) if self.y_entry.get() else 0
            wait_time = float(self.wait_entry.get()) if self.wait_entry.get() else 0
            description = self.desc_entry.get()
            
            step = {
                "type": step_type,
                "x": x,
                "y": y,
                "wait": wait_time,
                "desc": description
            }
            
            self.script_steps.append(step)
            self.update_steps_display()
            self.clear_inputs()
            self.set_status(f"已添加步骤: {description}")
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
    
    def clear_inputs(self):
        """清空输入框"""
        self.x_entry.delete(0, tk.END)
        self.y_entry.delete(0, tk.END)
        self.wait_entry.delete(0, tk.END)
        self.wait_entry.insert(0, "5")
        self.desc_entry.delete(0, tk.END)
    
    def delete_step(self):
        """删除选中的步骤"""
        selection = self.steps_tree.selection()
        if selection:
            index = int(selection[0]) - 1
            self.script_steps.pop(index)
            self.update_steps_display()
            self.set_status("已删除选中步骤")
        else:
            messagebox.showwarning("警告", "请先选择一个步骤")
    
    def update_steps_display(self):
        """更新步骤列表显示"""
        # 清空现有显示
        for item in self.steps_tree.get_children():
            self.steps_tree.delete(item)
        
        # 添加所有步骤
        for i, step in enumerate(self.script_steps, 1):
            self.steps_tree.insert("", "end", values=(
                i, step["type"], step["x"], step["y"], step["wait"], step["desc"]
            ))
    
    def execute_script(self):
        """执行脚本"""
        if not self.script_steps:
            messagebox.showwarning("警告", "没有可执行的步骤")
            return
        
        if self.is_playing:
            messagebox.showwarning("警告", "脚本正在执行中")
            return
        
        try:
            loop_count = int(self.loop_count.get())
            loop_interval = float(self.loop_interval.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的循环设置")
            return
        
        # 在后台线程中执行
        self.play_thread = threading.Thread(
            target=self.run_script, 
            args=(loop_count, loop_interval)
        )
        self.play_thread.daemon = True
        self.play_thread.start()
    
    def run_script(self, loop_count, loop_interval):
        """运行脚本的主逻辑"""
        self.is_playing = True
        self.play_btn.config(state="disabled")
        
        try:
            for loop in range(loop_count):
                if not self.is_playing:
                    break
                
                self.set_status(f"开始第 {loop + 1}/{loop_count} 次循环")
                
                for i, step in enumerate(self.script_steps):
                    if not self.is_playing:
                        break
                    
                    self.current_step_index = i
                    self.set_status(f"执行步骤 {i + 1}/{len(self.script_steps)}: {step['desc']}")
                    
                    # 执行步骤
                    if step["type"] == "点击":
                        pyautogui.click(step["x"], step["y"])
                        self.set_status(f"点击位置 ({step['x']}, {step['y']})")
                    elif step["type"] == "移动":
                        pyautogui.moveTo(step["x"], step["y"])
                        self.set_status(f"移动到 ({step['x']}, {step['y']})")
                    
                    # 等待指定时间
                    if step["wait"] > 0:
                        wait_remaining = step["wait"]
                        while wait_remaining > 0 and self.is_playing:
                            time.sleep(0.1)
                            wait_remaining -= 0.1
                
                # 循环间隔（除了最后一次）
                if loop < loop_count - 1 and self.is_playing:
                    self.set_status(f"等待 {loop_interval} 秒后开始下一次循环")
                    time.sleep(loop_interval)
            
            self.set_status("脚本执行完成")
            
        except Exception as e:
            self.set_status(f"执行出错: {str(e)}")
            messagebox.showerror("错误", f"执行过程中出现错误: {str(e)}")
        
        finally:
            self.is_playing = False
            self.root.after(0, self.enable_buttons)
    
    def stop_script(self):
        """停止脚本执行"""
        self.is_playing = False
        self.set_status("脚本执行已停止")
    
    def enable_buttons(self):
        """重新启用按钮"""
        self.play_btn.config(state="normal")
    
    def clear_script(self):
        """清除所有步骤"""
        if messagebox.askyesno("确认", "确定要清除所有步骤吗？"):
            self.script_steps.clear()
            self.update_steps_display()
            self.set_status("已清除所有步骤")
    
    def save_script(self):
        """保存脚本到文件"""
        if not self.script_steps:
            messagebox.showwarning("警告", "没有可保存的步骤")
            return
        
        filename = f"video_course_script_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump({
                    'timestamp': datetime.now().isoformat(),
                    'settings': {
                        'start_x': self.start_x.get(),
                        'start_y': self.start_y.get(),
                        'play_x': self.play_x.get(),
                        'play_y': self.play_y.get(),
                        'interval_x': self.interval_x.get(),
                        'interval_y': self.interval_y.get(),
                        'course_count': self.course_count.get(),
                        'video_duration': self.video_duration.get()
                    },
                    'steps': self.script_steps
                }, f, indent=2, ensure_ascii=False)
            
            self.set_status(f"脚本已保存到: {filename}")
            messagebox.showinfo("成功", f"脚本已保存到: {filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")
    
    def load_script(self):
        """从文件加载脚本"""
        from tkinter import filedialog
        
        filename = filedialog.askopenfilename(
            title="选择脚本文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.script_steps = data.get('steps', [])
                settings = data.get('settings', {})
                
                # 恢复设置
                for key, value in settings.items():
                    if hasattr(self, key):
                        getattr(self, key).delete(0, tk.END)
                        getattr(self, key).insert(0, str(value))
                
                self.update_steps_display()
                self.set_status(f"已加载脚本: {filename}")
                messagebox.showinfo("成功", f"已加载 {len(self.script_steps)} 个步骤")
                
            except Exception as e:
                messagebox.showerror("错误", f"加载失败: {str(e)}")
    
    def set_status(self, message):
        """设置状态栏消息"""
        self.status_label.config(text=f"{datetime.now().strftime('%H:%M:%S')} - {message}")

def main():
    
    root = tk.Tk()
    app = VideoCourseAutomator(root)
    root.mainloop()

if __name__ == "__main__":
    main()