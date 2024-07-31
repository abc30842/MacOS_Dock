import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, colorchooser
from PIL import Image, ImageTk
import os, win32api, win32con, win32gui, win32ui, subprocess, threading
import time, asyncio, json, pythoncom
from win32com.client import Dispatch

def make_command(filepath):
    return lambda: run_shortcut(filepath)

def run_shortcut(filepath):
    pythoncom.CoInitialize()  # 初始化 COM
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(filepath)
    target = shortcut.Targetpath
    subprocess.Popen(target)

class MacOSDock(tk.Tk): 
    def __init__(self): 
        super().__init__() 
        self.bg_color = 'gray'
        self.new_opacity = 1.0
        self.app_menu_showing = False
        self.set_up_gui() 
        self.attributes('-topmost', True)
        self.bind("<Button-1>", lambda e: self.set_app_menu_showing(False))
        self.shortcuts_dir = 'app_shortcuts'
        if not os.path.exists(self.shortcuts_dir):
            os.makedirs(self.shortcuts_dir)
        asyncio.run(self.load_settings())  # 使用 asyncio.run 來運行協程
        asyncio.run(self.load_app_shortcuts())  # 使用 asyncio.run 來運行協程

    async def load_app_shortcuts(self):
        for filename in os.listdir(self.shortcuts_dir):
            filepath = os.path.join(self.shortcuts_dir, filename)
            if filepath.endswith('.lnk'):  # 確保只處理捷徑檔案
                icon = await self.get_icon(filepath, original=True)  # 使用 await 來等待協程
                app_name = await self.get_app_name(filepath)  # 獲取應用程式名稱
                command = make_command(filepath)
                await self.add_app_to_dock(app_name, icon, command)
    
    async def get_app_name(self, filepath):
        pythoncom.CoInitialize()  # 初始化 COM
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(filepath)
        target = shortcut.Targetpath
        return os.path.basename(target)

    async def load_settings(self):
        if os.path.exists('settings.json'):
            with open('settings.json', 'r') as f:
                settings = json.load(f)
            self.bg_color = settings.get('bg_color', 'gray')
            self.new_opacity = settings.get('new_opacity', 1.0)
            self.configure(bg=self.bg_color)
            self.dock_frame.configure(bg=self.bg_color)
            self.attributes('-alpha', self.new_opacity)
            # 更新現有的新增按鈕的背景色
            for child in self.dock_frame.winfo_children():
                if isinstance(child, tk.Button) and child.cget('text') == '+':
                    child.configure(bg=self.bg_color, highlightthickness=0)
        await self.add_app_to_dock('新增', '+', self.add_app)
        # 確保新增按鈕的背景顏色也同步更新
        for child in self.dock_frame.winfo_children():
            if isinstance(child, tk.Button) and child.cget('text') == '新增':
                child.configure(bg=self.bg_color, highlightthickness=0)

    def save_settings(self):
        settings = {
            'bg_color': self.bg_color,
            'new_opacity': self.new_opacity,
            'shortcuts': [os.path.basename(f) for f in os.listdir(self.shortcuts_dir)]
        }
        with open('settings.json', 'w') as f:
            json.dump(settings, f)
        
    async def reset_settings(self):
        if os.path.exists('settings.json'):
            os.remove('settings.json')
        for widget in self.dock_frame.winfo_children():
            widget.destroy()
        self.bg_color = 'gray'
        self.new_opacity = 1.0
        self.configure(bg=self.bg_color)
        self.dock_frame.configure(bg=self.bg_color)
        self.attributes('-alpha', self.new_opacity)
        await self.add_app_to_dock('新增', '+', self.add_app)
        await self.load_app_shortcuts()

    async def get_icon(self, filepath, original):
        try:
            if original:
                # 檢查檔案是否為捷徑檔案
                if not filepath.lower().endswith(('.lnk', '.url')):
                    # 如果不是捷徑檔案，直接從應用程式抓取圖示
                    large, small = win32gui.ExtractIconEx(filepath, 0)
                    if large:
                        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                        hbmp = win32ui.CreateBitmap()
                        hbmp.CreateCompatibleBitmap(hdc, 32, 32)
                        hdc = hdc.CreateCompatibleDC()
                        hdc.SelectObject(hbmp)
                        hdc.DrawIcon((0, 0), large[0])
                        win32gui.DestroyIcon(large[0])
                        win32gui.DestroyIcon(small[0])
                        icon = ImageTk.PhotoImage(Image.frombuffer('RGBA', (32, 32), hbmp.GetBitmapBits(True), 'raw', 'BGRA', 0, 1))
                        hdc.DeleteDC()
                        return icon
                    else:
                        img = Image.new('RGB', (30, 30), color='white')
                        return ImageTk.PhotoImage(img)
                
                pythoncom.CoInitialize()  # 初始化 COM
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(filepath)
                target = shortcut.Targetpath
                large, small = win32gui.ExtractIconEx(target, 0)
                if large:
                    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                    hbmp = win32ui.CreateBitmap()
                    hbmp.CreateCompatibleBitmap(hdc, 32, 32)
                    hdc = hdc.CreateCompatibleDC()
                    hdc.SelectObject(hbmp)
                    hdc.DrawIcon((0, 0), large[0])
                    win32gui.DestroyIcon(large[0])
                    win32gui.DestroyIcon(small[0])
                    icon = ImageTk.PhotoImage(Image.frombuffer('RGBA', (32, 32), hbmp.GetBitmapBits(True), 'raw', 'BGRA', 0, 1))
                    hdc.DeleteDC()
                    return icon
                else:
                    img = Image.new('RGB', (30, 30), color='white')
                    return ImageTk.PhotoImage(img)
            else:
                img = Image.open(filepath)
                max_size = (30, 30)
                img.thumbnail(max_size)
                img.save(os.path.join(os.getcwd(), 'temp_icon.png'))
                return ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"Error loading icon: {e}")
            img = Image.new('RGB', (30, 30), color='white')
            return ImageTk.PhotoImage(img)
    
    def set_up_gui(self): 
        screen_width = self.winfo_screenwidth() 
        self.geometry(f'600x50+{int(screen_width/2)-300}+{self.winfo_screenheight()-85}') 
        self.overrideredirect(True) 
        self.configure(bg=self.bg_color) 
        self.attributes('-alpha', 0.5)
        self.dock_frame = tk.Frame(self, bg=self.bg_color) 
        self.dock_frame.pack(expand=True, fill='both') 

        self.right_click_menu = tk.Menu(self, tearoff=0)
        self.right_click_menu.add_command(label="新增", command=self.add_app)
        self.right_click_menu.add_command(label="改變背景顏色", command=self.change_bg_color)
        self.right_click_menu.add_command(label="改變透明度", command=self.change_opacity)
        self.right_click_menu.add_command(label="初始化設定", command=lambda: asyncio.run(self.reset_settings()))
        self.right_click_menu.add_command(label="關閉桌面小工具", command=self.destroy)
        self.bind("<Button-3>", self.show_right_click_menu)
        self.bind('<Enter>', self.show_dock)
        self.bind('<Leave>', self.hide_dock)
        threading.Thread(target=self.mouse_listener, daemon=True).start()

    def create_shortcut(self, target, shortcut_dir, name):
        shortcut_path = os.path.join(shortcut_dir, f"{name}.lnk")
        shell = Dispatch('WScript.Shell', pythoncom.CoInitialize())
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = os.path.dirname(target)
        shortcut.save()
        return shortcut_path

    def add_app(self):
        filepath = filedialog.askopenfilename(title="選擇應用程式")
        if filepath:
            asyncio.run(self.process_app_async(filepath))


    async def process_app_async(self, filepath):
        app_name = os.path.basename(filepath)
        shortcut_path = self.create_shortcut(filepath, self.shortcuts_dir, app_name)
        icon = await self.get_icon(filepath, original=True)  # 使用原始應用程式的圖示
        command = make_command(shortcut_path)
        await self.add_app_to_dock(app_name, icon, command)
        self.save_settings()

    async def add_app_to_dock(self, app_name, icon, command): 
        app_frame = tk.Frame(self.dock_frame, bg=self.bg_color)
        display_name = os.path.splitext(app_name)[0]
        
        if icon == '+':
            app_button = tk.Button(app_frame, text=icon, command=command, bg=self.bg_color, fg='white', bd=0, highlightthickness=0)
        else:
            app_button = tk.Button(app_frame, image=icon, command=command, bg=self.bg_color, bd=0, highlightthickness=0) 
            app_button.image = icon 
            app_button.bind('<Enter>', lambda e: app_button.config(image=icon))
            app_button.bind('<Leave>', lambda e: app_button.config(image=icon))

        menu = tk.Menu(app_button, tearoff=0)
        menu.add_command(label="修改名稱", command=lambda: [self.rename_app(app_button, app_label), self.set_app_menu_showing(False)])
        menu.add_command(label="修改圖示", command=lambda: [self.change_icon(app_button), self.set_app_menu_showing(False)])
        app_button.bind("<Button-3>", lambda e: [menu.post(e.x_root, e.y_root), self.set_app_menu_showing(True)])
        menu.add_command(label="刪除", command=lambda: [self.remove_app(app_frame, app_button, app_label), self.set_app_menu_showing(False)])

        app_button.pack()
        app_button.fullname = app_name

        display_name = display_name[:10] + '...' if len(display_name) > 10 else display_name

        app_label = tk.Label(app_frame, text=display_name, bg=self.bg_color, fg='white')
        app_label.bind('<Enter>', lambda e: app_label.config(text=app_name))
        app_label.bind('<Leave>', lambda e: app_label.config(text=display_name))

        app_label.pack()
        app_frame.pack(side='left', padx=5)

    def rename_app(self, app_button, app_label):
        new_name = simpledialog.askstring("修改名稱", "請輸入新的名稱")
        if new_name:
            app_label.config(text=new_name)
            app_label.bind('<Enter>', lambda e: app_label.config(text=new_name))
            app_label.bind('<Leave>', lambda e: app_label.config(text=new_name))

    def remove_app(self, app_frame, app_button, app_label):
        app_name = app_button.fullname
        shortcut_path = os.path.join(self.shortcuts_dir, app_name + '.lnk')  # 確保捷徑檔案名稱正確
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
        app_frame.destroy()
        app_button.destroy()
        app_label.destroy()

    def set_app_menu_showing(self, is_showing):
        self.app_menu_showing = is_showing

    def show_right_click_menu(self, event):
        if self.app_menu_showing:
            return
        self.right_click_menu.post(event.x_root, event.y_root)
        self.set_app_menu_showing(False)
        
    def change_bg_color(self):
        new_color = colorchooser.askcolor()[1]
        if new_color:
            self.bg_color = new_color
            self.configure(bg=self.bg_color)
            self.dock_frame.configure(bg=self.bg_color)
            for child in self.dock_frame.winfo_children():
                child.configure(bg=self.bg_color)
                for grandchild in child.winfo_children():
                    grandchild.configure(bg=self.bg_color)
            self.save_settings()  # 儲存設定
    
    def change_icon(self, app_button):
        iconpath = filedialog.askopenfilename(title="選擇圖示檔案")
        if iconpath:
            # 使用 asyncio.run 來執行協程並獲取返回值
            normal_icon = asyncio.run(self.get_icon(iconpath, original=False))
            app_button.config(image=normal_icon)
            app_button.image = normal_icon

            # 確保圖示對象不會被垃圾回收
            if not hasattr(self, 'icons'):
                self.icons = []
            self.icons.append(normal_icon)

            # 在滑鼠進入和離開時切換圖示
            app_button.bind('<Enter>', lambda e: app_button.config(image=normal_icon))
            app_button.bind('<Leave>', lambda e: app_button.config(image=normal_icon))

    def change_opacity(self):
        opacity_window = tk.Toplevel(self)
        opacity_window.title("改變透明度")
        opacity_window.geometry("250x100+300+300")

        opacity_label = ttk.Label(opacity_window, text=f"透明度: {self.new_opacity * 100}%")
        opacity_label.pack()

        opacity_scale = ttk.Scale(opacity_window, from_=0, to=100, orient='horizontal', length=200, command=lambda value: opacity_label.config(text=f"當前透明度: {round(float(value), 1)}%"))
        opacity_scale.set(self.new_opacity * 100)
        opacity_scale.pack()

        confirm_button = ttk.Button(opacity_window, text="確認", command=lambda: [self.update_opacity(opacity_scale.get()), opacity_window.destroy()])
        confirm_button.pack()

    def update_opacity(self, value):
        self.new_opacity = float(value) / 100
        self.attributes('-alpha', self.new_opacity)
        self.save_settings()  # 儲存設定

    def show_dock(self, event=None):
        # 獲取滑鼠的當前位置
        x, y = self.winfo_pointerxy()

        # 檢查滑鼠是否在 dock 的上方
        if self.winfo_rooty() - y <= 100:
            self.attributes('-alpha', self.new_opacity)  # 顯示 dock
        else:
            self.after(100, self.show_dock)  # 100 毫秒後再次檢查滑鼠的位置

    def hide_dock(self, event=None):
        # 獲取滑鼠的當前位置
        x, y = self.winfo_pointerxy()

        # 檢查滑鼠是否在 dock 的上方
        if self.winfo_rooty() - y > 100:
            self.attributes('-alpha', 0.0)  # 隱藏 dock
        else:
            self.after(100, self.hide_dock)  # 100 毫秒後再次檢查滑鼠的位置

    def mouse_listener(self):
        while True:
            # 獲取滑鼠的當前位置
            x, y = win32api.GetCursorPos()

            # 獲取 dock 的位置和大小
            dock_x = self.winfo_rootx()
            dock_y = self.winfo_rooty()
            dock_width = self.winfo_width()
            dock_height = self.winfo_height()

            # 檢查滑鼠是否在 dock 的範圍內
            if dock_x <= x <= dock_x + dock_width and dock_y <= y <= dock_y + dock_height:
                self.after(0, self.show_dock)
            else:
                self.after(0, self.hide_dock)

            # 等待一段時間以減少 CPU 使用率
            time.sleep(0.1)

if __name__ == '__main__':
    app = MacOSDock()
    app.mainloop()
