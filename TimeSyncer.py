import ntplib
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from win32com.client import Dispatch
import winshell
import ctypes
import configparser
import pystray
from PIL import Image, ImageTk
from ttkthemes import ThemedTk
import threading

# 获取 icon.ico 的正确路径
if getattr(sys, 'frozen', False):  # 是否为打包后的可执行文件
    base_path = sys._MEIPASS  # 打包后的工作目录
else:
    base_path = os.path.dirname(os.path.abspath(__file__))  # 未打包时的工作目录

icon_path = os.path.join(base_path, "icon.ico")

# 默认使用的 NTP 服务器
DEFAULT_NTP_SERVERS = [
    "time.windows.com",
    "pool.ntp.org",
    "ntp.aliyun.com",
    "time.apple.com",
    "time.google.com"
]

# 配置文件路径
CONFIG_FILE = "config.ini"

def is_admin():
    """检查是否以管理员身份运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员身份重新运行程序"""
    if os.name == 'nt':
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

class TimeSynchronizer:
    def __init__(self, root):
        self.root = root
        self.root.title("时间同步 1.0 --QwejayHuang")
        self.root.geometry("400x400")

        # 检查 icon.ico 文件是否存在
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        else:
            print("Warning: Icon file 'icon.ico' not found.")

        # 加载配置
        self.config = self.load_config()

        # 初始化界面
        self.setup_ui()

        # 初始化复选框状态
        self.auto_start_var = tk.BooleanVar(value=self.is_auto_start_enabled())
        self.auto_start_checkbox.config(variable=self.auto_start_var)

        # 初始化隐藏主界面复选框状态
        self.hide_on_start_var = tk.BooleanVar(value=self.config.getboolean('Settings', 'hide_on_start', fallback=False))
        self.hide_on_start_checkbox.config(variable=self.hide_on_start_var)

        # 如果设置为启动时隐藏主界面，则隐藏主界面并显示托盘图标
        if self.hide_on_start_var.get():
            self.hide_main_window()
            self.auto_sync_time()

        # 设置窗口图标
        if os.path.exists(icon_path):
            icon_image = Image.open(icon_path)
            self.icon_photo = ImageTk.PhotoImage(icon_image)
            self.root.iconphoto(False, self.icon_photo)

    def setup_ui(self):
        """设置GUI界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title label
        title_label = ttk.Label(main_frame, text="系统时间同步工具", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=10)

        # Sync button
        sync_button = ttk.Button(main_frame, text="立即同步时间", command=self.sync_time, style="Accent.TButton")
        sync_button.pack(pady=10, fill=tk.X)

        # 设置区域框架
        settings_frame = ttk.LabelFrame(main_frame, text="设置", padding="10")
        settings_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Auto-start checkbox
        self.auto_start_checkbox = ttk.Checkbutton(
            settings_frame,
            text="开机自启动",
            style="Toggle.TCheckbutton",
            command=self.on_auto_start_toggle,
        )
        self.auto_start_checkbox.pack(pady=5, anchor=tk.W)

        # Hide on start checkbox
        self.hide_on_start_checkbox = ttk.Checkbutton(
            settings_frame,
            text="启动时隐藏主界面",
            style="Toggle.TCheckbutton",
            command=self.on_hide_on_start_toggle,
        )
        self.hide_on_start_checkbox.pack(pady=5, anchor=tk.W)

        # NTP服务器设置按钮
        ntp_button = ttk.Button(settings_frame, text="设置NTP服务器", command=self.open_ntp_settings)
        ntp_button.pack(pady=10, fill=tk.X)

        # 状态栏框架
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Status label
        self.status_label = ttk.Label(status_frame, text="点击“立即同步时间”以更新时间。", font=("微软雅黑", 9))
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 监听窗口的最小化事件
        self.root.protocol("WM_ICONIFY", self.on_minimize)

    def open_ntp_settings(self):
        """打开NTP服务器设置界面"""
        self.ntp_settings_window = tk.Toplevel(self.root)
        self.ntp_settings_window.title("NTP服务器设置")
        self.ntp_settings_window.geometry("500x400")

        # 主框架
        main_frame = ttk.Frame(self.ntp_settings_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # NTP服务器列表框
        self.ntp_listbox = tk.Listbox(main_frame, selectmode=tk.SINGLE, font=("微软雅黑", 9))
        self.ntp_listbox.pack(fill=tk.BOTH, expand=True, pady=10)

        # 添加默认服务器到列表框
        for server in self.get_ntp_servers():
            self.ntp_listbox.insert(tk.END, server)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        # 添加按钮
        add_button = ttk.Button(button_frame, text="添加服务器", command=self.add_ntp_server)
        add_button.pack(side=tk.LEFT, padx=5, expand=True)

        # 删除按钮
        remove_button = ttk.Button(button_frame, text="删除服务器", command=self.remove_ntp_server)
        remove_button.pack(side=tk.LEFT, padx=5, expand=True)

        # 上移按钮
        up_button = ttk.Button(button_frame, text="上移", command=self.move_up)
        up_button.pack(side=tk.LEFT, padx=5, expand=True)

        # 下移按钮
        down_button = ttk.Button(button_frame, text="下移", command=self.move_down)
        down_button.pack(side=tk.LEFT, padx=5, expand=True)

    def add_ntp_server(self):
        """添加新的NTP服务器"""
        new_server = simpledialog.askstring("添加NTP服务器", "请输入新的NTP服务器地址：", parent=self.ntp_settings_window)
        if new_server and new_server not in self.get_ntp_servers():
            servers = self.get_ntp_servers()
            servers.append(new_server)
            self.config.set('NTP', 'servers', ','.join(servers))
            self.save_config()
            self.ntp_listbox.insert(tk.END, new_server)
            self.show_status(f"已添加服务器: {new_server}", fg="green")
        elif new_server in self.get_ntp_servers():
            self.show_status("该服务器已存在。", fg="red")

    def remove_ntp_server(self):
        """删除选中的NTP服务器"""
        try:
            selected_index = self.ntp_listbox.curselection()[0]
            remove_server = self.ntp_listbox.get(selected_index)
            servers = self.get_ntp_servers()
            servers.pop(selected_index)
            self.config.set('NTP', 'servers', ','.join(servers))
            self.save_config()
            self.ntp_listbox.delete(selected_index)
            self.show_status(f"已删除服务器: {remove_server}", fg="green")
        except IndexError:
            self.show_status("未选择任何服务器。", fg="red")

    def move_up(self):
        """上移选中的NTP服务器"""
        try:
            selected_index = self.ntp_listbox.curselection()[0]
            if selected_index > 0:
                servers = self.get_ntp_servers()
                server = servers.pop(selected_index)
                servers.insert(selected_index - 1, server)
                self.config.set('NTP', 'servers', ','.join(servers))
                self.save_config()
                self.ntp_listbox.delete(selected_index)
                self.ntp_listbox.insert(selected_index - 1, server)
                self.ntp_listbox.select_set(selected_index - 1)
        except IndexError:
            self.show_status("未选择任何服务器。", fg="red")

    def move_down(self):
        """下移选中的NTP服务器"""
        try:
            selected_index = self.ntp_listbox.curselection()[0]
            if selected_index < len(self.get_ntp_servers()) - 1:
                servers = self.get_ntp_servers()
                server = servers.pop(selected_index)
                servers.insert(selected_index + 1, server)
                self.config.set('NTP', 'servers', ','.join(servers))
                self.save_config()
                self.ntp_listbox.delete(selected_index)
                self.ntp_listbox.insert(selected_index + 1, server)
                self.ntp_listbox.select_set(selected_index + 1)
        except IndexError:
            self.show_status("未选择任何服务器。", fg="red")

    def get_startup_folder(self):
        """获取启动文件夹路径"""
        return os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')

    def is_auto_start_enabled(self):
        """检查是否已设置开机自启动"""
        shortcut_path = os.path.join(self.get_startup_folder(), 'TimeSync.lnk')
        return os.path.exists(shortcut_path)

    def set_auto_start(self, enabled):
        """设置或取消开机自启动"""
        shortcut_path = os.path.join(self.get_startup_folder(), 'TimeSync.lnk')
        try:
            if enabled:
                if not os.path.exists(shortcut_path):
                    shell = Dispatch('WScript.Shell')
                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.Targetpath = sys.executable  # Python解释器路径
                    shortcut.Arguments = f'"{os.path.abspath(__file__)}"'  # 脚本路径
                    shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
                    shortcut.save()
                    self.show_status("开机自启动已启用。", fg="green")
            else:
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                    self.show_status("程序已取消开机自启动。", fg="green")
        except Exception as e:
            self.show_status(f"设置开机自启动失败: {e}", fg="red")
        self.config.set('Settings', 'auto_start', str(enabled))
        self.save_config()

    def on_auto_start_toggle(self):
        """当复选框状态改变时自动执行"""
        auto_start = self.auto_start_var.get()
        self.set_auto_start(auto_start)

    def on_hide_on_start_toggle(self):
        """当隐藏主界面复选框状态改变时自动执行"""
        hide_on_start = self.hide_on_start_var.get()
        self.config.set('Settings', 'hide_on_start', str(hide_on_start))
        self.save_config()

    def get_network_time(self):
        """从默认的 NTP 服务器获取网络时间"""
        for server in self.get_ntp_servers():
            try:
                ntp_client = ntplib.NTPClient()
                response = ntp_client.request(server, timeout=3)  # 设置超时时间为3秒
                return datetime.fromtimestamp(response.tx_time)
            except Exception as e:
                continue
        return None

    def set_system_time(self, new_time):
        """设置系统时间"""
        try:
            time_str = new_time.strftime('%Y-%m-%d %H:%M:%S')
            os.system(f'date {time_str}')
            self.show_status(f"系统时间已设置为: {time_str}", fg="green")
        except Exception as e:
            self.show_status(f"设置系统时间失败: {e}", fg="red")

    def sync_time(self):
        """同步时间"""
        network_time = self.get_network_time()
        if network_time:
            self.set_system_time(network_time)
        else:
            self.show_status("无法从NTP服务器获取时间。", fg="red")

    def show_status(self, message, fg="black"):
        """在状态栏显示消息"""
        self.status_label.config(text=message)

    def load_config(self):
        """加载配置文件"""
        config = configparser.ConfigParser()
        if os.path.exists(CONFIG_FILE):
            config.read(CONFIG_FILE)
        else:
            # 初始化默认配置
            config['Settings'] = {'auto_start': 'False', 'hide_on_start': 'False'}
            config['NTP'] = {'servers': ','.join(DEFAULT_NTP_SERVERS)}
            with open(CONFIG_FILE, 'w') as f:
                config.write(f)
        return config

    def save_config(self):
        """保存配置文件"""
        with open(CONFIG_FILE, 'w') as f:
            self.config.write(f)

    def get_ntp_servers(self):
        """获取NTP服务器列表"""
        return self.config.get('NTP', 'servers', fallback=','.join(DEFAULT_NTP_SERVERS)).split(',')

    def on_minimize(self):
        """当窗口最小化时触发"""
        if self.hide_on_start_var.get():
            # 如果勾选了“启动时隐藏主界面”，则隐藏窗口并显示托盘图标
            self.hide_main_window()
        else:
            # 否则正常最小化到任务栏
            self.root.iconify()

    def hide_main_window(self):
        """隐藏主界面并显示托盘图标"""
        self.root.withdraw()  # 隐藏主窗口
        # 在独立线程中运行托盘图标
        threading.Thread(target=self.create_system_tray_icon, daemon=True).start()
        # 自动同步时间
        self.auto_sync_time()

    def show_main_window(self):
        """显示主界面"""
        self.icon.stop()  # 停止托盘图标
        self.root.deiconify()  # 恢复窗口
        self.root.lift()  # 将窗口置于最前面

    def auto_sync_time(self):
        """自动同步时间"""
        self.sync_time()

    def create_system_tray_icon(self):
        """创建系统托盘图标"""
        # 检查托盘图标文件是否存在
        if os.path.exists(icon_path):
            image = Image.open(icon_path)
        else:
            image = None

        menu = pystray.Menu(
            pystray.MenuItem("显示主界面", self.show_main_window),
            pystray.MenuItem("退出", self.quit_application)
        )
        self.icon = pystray.Icon("TimeSync", image, "时间同步工具", menu)
        self.icon.run()  # 运行托盘图标

    def quit_application(self):
        """退出应用程序"""
        if hasattr(self, 'icon'):
            self.icon.stop()  # 停止托盘图标
        self.root.quit()  # 退出主程序

    def check_single_instance(self):
        """检查是否已有实例在运行"""
        mutex_name = "Global\\TimeSyncMutex"
        self.mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
        if ctypes.windll.kernel32.GetLastError() == 183:
            # 已有实例在运行
            self.show_main_window()
            self.quit_application()
        else:
            # 没有实例在运行
            self.mutex = self.mutex

    def run(self):
        """运行程序"""
        self.check_single_instance()
        # 如果设置为启动时隐藏主界面，则自动同步时间
        if self.hide_on_start_var.get():
            self.auto_sync_time()
        self.root.mainloop()

if __name__ == "__main__":
    # 检查是否以管理员身份运行
    if os.name == 'nt' and not is_admin():
        run_as_admin()
    else:
        # 启动GUI
        root = ThemedTk(theme="superhero")
        app = TimeSynchronizer(root)
        app.run()