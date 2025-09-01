import pandas as pd
import os
from datetime import datetime, timedelta
from icalendar import Calendar, Event
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import hashlib
import sys

class XlsxToIcsConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel课表转ICS工具")
        self.root.geometry("750x750")  # 增加窗口高度
        self.root.resizable(True, True)  # 允许调整窗口大小
        self.root.configure(bg='#f0f0f0')
        
        # 存储选择的文件
        self.selected_files = []
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup user interface"""
        # Main frame
        main_frame = tk.Frame(self.root, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection button
        select_btn = tk.Button(main_frame, text="Select Excel Files", command=self.select_files,
                              bg='#3498db', fg='white', font=('Arial', 12, 'bold'),
                              relief='flat', padx=30, pady=8)
        select_btn.pack(pady=10)
        
        # File list frame
        list_frame = tk.Frame(main_frame, bg='#f0f0f0')
        list_frame.pack(fill=tk.BOTH, expand=True, pady=15)
        
        tk.Label(list_frame, text="Selected Files:", font=('Arial', 12, 'bold'), 
                bg='#f0f0f0', fg='#000000').pack(anchor='w')  # Changed to black for visibility
        
        # Create listbox with scrollbar
        list_container = tk.Frame(list_frame, bg='#ffffff', relief='sunken', bd=2)
        list_container.pack(fill=tk.BOTH, expand=True, pady=8)
        
        scrollbar = tk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(list_container, yscrollcommand=scrollbar.set, 
                                      font=('Arial', 10), height=8, 
                                      selectmode=tk.EXTENDED,
                                      bg='#ffffff', fg='#000000',  # Changed to black for visibility
                                      selectbackground='#3498db',
                                      selectforeground='white')
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Setup file operations
        self.setup_drag_drop(list_container)
        
        # Deduplication option
        self.enable_dedup = tk.BooleanVar(value=True)
        dedup_check = tk.Checkbutton(main_frame, text="Enable deduplication", 
                                    variable=self.enable_dedup, font=('Arial', 10),
                                    bg='#f0f0f0', fg='#000000')  # Changed to black for visibility
        dedup_check.pack(pady=10)
        
        # Convert button
        convert_btn = tk.Button(main_frame, text="Convert to ICS", command=self.convert_to_ics,
                               bg='#e74c3c', fg='white', font=('Arial', 12, 'bold'),
                               relief='flat', padx=30, pady=8)
        convert_btn.pack(pady=15)
        
        # Status label
        self.status_label = tk.Label(main_frame, text="Select Excel files to start",
                                    font=('Arial', 10), fg='#000000', bg='#f0f0f0')  # Changed to black for visibility
        self.status_label.pack(pady=10)
        
    
    def setup_drag_drop(self, widget):
        """设置文件操作功能"""
        # 添加双击选择文件功能
        widget.bind('<Double-Button-1>', lambda e: self.select_files())
        self.file_listbox.bind('<Double-Button-1>', lambda e: self.select_files())
        
        # 添加右键菜单
        self.setup_context_menu()
    
    def setup_context_menu(self):
        """设置右键菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="📁 选择文件", command=self.select_files)
        self.context_menu.add_command(label="🗑️ 移除选中", command=self.remove_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🧹 清除全部", command=self.clear_files)
        
        # 绑定右键菜单到列表框
        self.file_listbox.bind('<Button-2>', self.show_context_menu)  # Mac右键
        self.file_listbox.bind('<Control-Button-1>', self.show_context_menu)  # Mac Ctrl+点击
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def select_files(self):
        """选择Excel文件"""
        files = filedialog.askopenfilenames(
            title="选择Excel课表文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        added_count = 0
        for file_path in files:
            if file_path not in self.selected_files:
                self.selected_files.append(file_path)
                filename = os.path.basename(file_path)
                self.file_listbox.insert(tk.END, f"📄 {filename}")
                added_count += 1
        
        if added_count > 0:
            self.update_status(f"✅ 成功添加 {added_count} 个文件，共 {len(self.selected_files)} 个文件")
        else:
            self.update_status(f"📊 当前共 {len(self.selected_files)} 个文件")
    
    def remove_selected(self):
        """移除选中的文件"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("⚠️ 警告", "请先选择要移除的文件")
            return
        
        # 从后往前删除，避免索引变化
        for index in reversed(selected_indices):
            del self.selected_files[index]
            self.file_listbox.delete(index)
        
        self.update_status(f"🗑️ 已移除 {len(selected_indices)} 个文件")
    
    def clear_files(self):
        """清除文件列表"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_status("🧹 文件列表已清空")
    
    def update_status(self, message):
        """更新状态信息"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def generate_event_hash(self, row):
        """生成事件的唯一哈希值用于去重"""
        hash_string = f"{row['日期']}_{row['開始時間']}_{row['結束時間']}_{row['科目名稱']}_{row['課室']}"
        return hashlib.md5(hash_string.encode('utf-8')).hexdigest()
    
    def process_excel_data(self, df):
        """处理Excel数据并返回事件列表"""
        events = []
        
        for index, row in df.iterrows():
            try:
                # 解析时间
                date = pd.to_datetime(str(row['日期'])).date()
                start_time = pd.to_datetime(str(row['開始時間'])).time()
                end_time = pd.to_datetime(str(row['結束時間'])).time()
                
                # 生成事件的开始和结束时间
                start_datetime = datetime.combine(date, start_time)
                end_datetime = datetime.combine(date, end_time)
                
                # 创建事件数据
                event_data = {
                    'summary': f"{row['科目名稱']} ({row['班別名稱']})",
                    'dtstart': start_datetime,
                    'dtend': end_datetime,
                    'location': str(row['課室']),
                    'description': f"教师: {row['教師']}",
                    'hash': self.generate_event_hash(row)
                }
                
                events.append(event_data)
                
            except Exception as e:
                print(f"处理行 {index} 时出错: {e}")
                continue
        
        return events
    
    def remove_duplicates(self, events):
        """去除重复事件"""
        seen_hashes = set()
        unique_events = []
        
        for event in events:
            if event['hash'] not in seen_hashes:
                seen_hashes.add(event['hash'])
                unique_events.append(event)
        
        return unique_events
    
    def convert_to_ics(self):
        """转换Excel文件为ICS格式"""
        if not self.selected_files:
            messagebox.showwarning("⚠️ 警告", "请先选择Excel文件")
            return
        
        try:
            self.update_status("🔄 正在转换...")
            
            # 创建日历对象
            cal = Calendar()
            cal.add('prodid', '-//Excel To ICS//mxm.dk//')
            cal.add('version', '2.0')
            
            all_events = []
            processed_files = 0
            
            # 处理每个文件
            for file_path in self.selected_files:
                try:
                    filename = os.path.basename(file_path)
                    self.update_status(f"📖 正在处理: {filename}")
                    df = pd.read_excel(file_path)
                    events = self.process_excel_data(df)
                    all_events.extend(events)
                    processed_files += 1
                    self.update_status(f"✅ 已处理 {processed_files}/{len(self.selected_files)} 个文件")
                except Exception as e:
                    messagebox.showerror("❌ 错误", f"处理文件 {os.path.basename(file_path)} 时出错:\n{str(e)}")
                    continue
            
            if not all_events:
                messagebox.showwarning("⚠️ 警告", "没有找到有效的课程数据")
                self.update_status("❌ 转换失败 - 无有效数据")
                return
            
            # 去重处理
            removed_count = 0
            if self.enable_dedup.get():
                self.update_status("🔍 正在去重...")
                original_count = len(all_events)
                all_events = self.remove_duplicates(all_events)
                removed_count = original_count - len(all_events)
                if removed_count > 0:
                    self.update_status(f"🔄 已去除 {removed_count} 个重复事件")
            
            # 添加事件到日历
            self.update_status("📝 正在生成ICS文件...")
            for event_data in all_events:
                event = Event()
                event.add('summary', event_data['summary'])
                event.add('dtstart', event_data['dtstart'])
                event.add('dtend', event_data['dtend'])
                if event_data.get('description'):
                    event.add('description', event_data['description'])
                if event_data.get('location'):
                    event.add('location', event_data['location'])
                event.add('uid', f"{event_data['hash']}@xlsxtoics.local")
                cal.add_component(event)
            
            # 保存文件
            downloads_dir = os.path.expanduser("~/Downloads")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(downloads_dir, f"课程表_{timestamp}.ics")
            
            with open(output_file, 'wb') as f:
                f.write(cal.to_ical())
            
            # 显示成功信息
            success_msg = (f"🎉 转换完成！\n\n"
                          f"📁 保存位置: {output_file}\n"
                          f"📊 处理文件: {processed_files} 个\n"
                          f"📅 生成事件: {len(all_events)} 个")
            
            if self.enable_dedup.get() and removed_count > 0:
                success_msg += f"\n🔄 去重事件: {removed_count} 个"
            
            success_msg += f"\n\n💡 现在可以将此文件导入到您的日历应用中！"
            
            messagebox.showinfo("✅ 转换成功", success_msg)
            self.update_status(f"🎉 转换完成 - 共生成 {len(all_events)} 个事件")
            
        except Exception as e:
            messagebox.showerror("❌ 错误", f"转换过程中出现错误:\n{str(e)}")
            self.update_status("❌ 转换失败")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

def main():
    app = XlsxToIcsConverter()
    app.run()

if __name__ == "__main__":
    main()