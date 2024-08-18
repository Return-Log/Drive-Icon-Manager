import sys
import winreg
import os
import ctypes
import time
import pyautogui
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QMessageBox, QTabWidget, QListWidget, QListWidgetItem
from PyQt6.QtCore import Qt

# 定义版本号和链接
VERSION = "v2.0"
GITHUB_LINK = "https://github.com/Return-Log/Drive-Icon-Manager"
FORUM_LINK = "https://www.52pojie.cn/home.php?mod=space&uid=2286792"

class DriveIconManager(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.icons = []
        self.selected_location = '此电脑'  # 默认选择“此电脑”
        self.display_icons()

    def initUI(self):
        self.setWindowTitle('Drive Icon Manager')
        self.setGeometry(300, 300, 500, 400)

        layout = QVBoxLayout()

        # 欢迎和链接标签
        self.welcome_label = QLabel(f"欢迎使用 Drive Icon Manager 版本 {VERSION}", self)
        layout.addWidget(self.welcome_label)

        # 链接布局
        link_layout = QHBoxLayout()
        self.github_link = QLabel(f"<a href='{GITHUB_LINK}'>访问 GitHub 仓库</a>", self)
        self.github_link.setOpenExternalLinks(True)
        link_layout.addWidget(self.github_link)

        self.forum_link = QLabel(f"<a href='{FORUM_LINK}'>访问 吾爱破解论坛</a>", self)
        self.forum_link.setOpenExternalLinks(True)
        link_layout.addWidget(self.forum_link)

        layout.addLayout(link_layout)

        # 不同位置的 Tab Widget
        self.tab_widget = QTabWidget(self)
        self.this_pc_tab = QWidget()
        self.sidebar_tab = QWidget()

        self.tab_widget.addTab(self.this_pc_tab, "此电脑")
        self.tab_widget.addTab(self.sidebar_tab, "资源管理器侧边栏")
        self.tab_widget.currentChanged.connect(self.on_tab_change)

        # 设置标签页布局
        self.this_pc_layout = QVBoxLayout()
        self.this_pc_text = QListWidget()
        self.this_pc_layout.addWidget(self.this_pc_text)
        self.this_pc_tab.setLayout(self.this_pc_layout)

        self.sidebar_layout = QVBoxLayout()
        self.sidebar_text = QListWidget()
        self.sidebar_layout.addWidget(self.sidebar_text)
        self.sidebar_tab.setLayout(self.sidebar_layout)

        # 向布局添加标签
        layout.addWidget(self.tab_widget)

        # 刷新和删除按钮的水平布局
        button_layout = QHBoxLayout()

        self.refresh_button = QPushButton('刷新', self)
        self.refresh_button.clicked.connect(self.display_icons)
        self.refresh_button.setFixedWidth(60)
        button_layout.addWidget(self.refresh_button)

        self.delete_button = QPushButton('删除选中的驱动器图标', self)
        self.delete_button.clicked.connect(self.delete_selected_icon)
        button_layout.addWidget(self.delete_button)

        self.permissions_button = QPushButton('更改注册表权限', self)
        self.permissions_button.clicked.connect(self.open_permissions_window)
        button_layout.addWidget(self.permissions_button)

        self.exit_button = QPushButton('退出程序', self)
        self.exit_button.clicked.connect(self.close)
        button_layout.addWidget(self.exit_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

        # 以管理员身份运行 check
        self.run_as_admin()

    def is_admin(self):
        """检查是否具有管理员权限"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def run_as_admin(self):
        """以管理员权限重新运行"""
        if not self.is_admin():
            QMessageBox.critical(self, "权限不足", "本工具需要管理员权限，请重新运行并获取管理员权限")
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
            sys.exit()

    def get_current_user_sid(self):
        """通过注册表获取当前用户的SID"""
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI")
            sid, _ = winreg.QueryValueEx(key, "LastLoggedOnUserSid")
            winreg.CloseKey(key)
            return sid
        except Exception as e:
            self.display_error_message(f"无法获取当前用户的SID: {e}")
            return None

    def list_drive_icons(self, base_key, path, source):
        """列出指定路径下的驱动器图标及其注册表项名称和显示名称"""
        icons = []
        try:
            key = winreg.OpenKey(base_key, path, 0, winreg.KEY_READ)
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkey_path = f"{path}\\{subkey_name}"
                    subkey = winreg.OpenKey(base_key, subkey_path, 0, winreg.KEY_READ)
                    try:
                        display_name, _ = winreg.QueryValueEx(subkey, None)
                    except FileNotFoundError:
                        display_name = "无显示名称"
                    icons.append((subkey_name, display_name, source))
                    winreg.CloseKey(subkey)
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except Exception as e:
            self.display_error_message(f"无法列出驱动器图标: {e}")
        return icons

    def delete_drive_icon(self, index):
        """删除指定索引的驱动器图标"""
        subkey_name, _, source = self.icons[index]
        if source == '此电脑':
            key_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
        else:
            current_user_sid = self.get_current_user_sid()
            key_path = fr'{current_user_sid}\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER if source == '此电脑' else winreg.HKEY_USERS, key_path, 0,
                                 winreg.KEY_ALL_ACCESS)
            winreg.DeleteKey(key, subkey_name)
            winreg.CloseKey(key)
            QMessageBox.information(self, "成功", f"已删除驱动器图标: {subkey_name}")
        except Exception as e:
            self.display_error_message(f"无法删除驱动器图标: {e}")
        finally:
            self.display_icons()  # 删除后自动刷新

    def display_icons(self):
        """显示所有驱动器图标"""
        self.icons = []
        if self.selected_location == '此电脑':
            self.this_pc_text.clear()
            self.icons += self.list_drive_icons(winreg.HKEY_CURRENT_USER,
                                                r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace',
                                                '此电脑')
            if self.icons:
                for index, (subkey_name, display_name, source) in enumerate(self.icons):
                    item = QListWidgetItem(f"{subkey_name} - {display_name}")
                    self.this_pc_text.addItem(item)
            else:
                item = QListWidgetItem("此电脑未找到驱动器图标")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsSelectable)  # 设置不可选择
                self.this_pc_text.addItem(item)

        elif self.selected_location == '资源管理器侧边栏':
            self.sidebar_text.clear()
            current_user_sid = self.get_current_user_sid()
            self.icons += self.list_drive_icons(winreg.HKEY_USERS,
                                                fr'{current_user_sid}\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace',
                                                '资源管理器侧边栏')
            if self.icons:
                for index, (subkey_name, display_name, source) in enumerate(self.icons):
                    item = QListWidgetItem(f"{subkey_name} - {display_name}")
                    self.sidebar_text.addItem(item)
            else:
                item = QListWidgetItem("资源管理器侧边栏未找到驱动器图标")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsSelectable)  # 设置不可选择
                self.sidebar_text.addItem(item)

    def delete_selected_icon(self):
        """删除选中的驱动器图标"""
        if self.selected_location == '此电脑':
            selected_index = self.this_pc_text.currentRow()
        else:
            selected_index = self.sidebar_text.currentRow()

        if selected_index >= 0:
            reply = QMessageBox.question(self, '确认删除',
                                         "你确定要删除这个驱动器图标吗?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)

            if reply == QMessageBox.StandardButton.Yes:
                self.delete_drive_icon(selected_index)

    def display_error_message(self, message):
        """在当前标签页显示错误信息"""
        item = QListWidgetItem(message)
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsSelectable)  # 设置不可选择
        if self.selected_location == '此电脑':
            self.this_pc_text.addItem(item)
        elif self.selected_location == '资源管理器侧边栏':
            self.sidebar_text.addItem(item)

    def open_permissions_window(self):
        """打开注册表namespace文件夹权限设置窗口"""
        if self.selected_location == '此电脑':
            key_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
            full_key_path = f"HKEY_CURRENT_USER\\{key_path}"
        elif self.selected_location == '资源管理器侧边栏':
            current_user_sid = self.get_current_user_sid()
            key_path = fr"{current_user_sid}\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace"
            full_key_path = f"HKEY_USERS\\{key_path}"

        # 打开注册表编辑器
        os.system('start regedit')

        time.sleep(1)  # 等待注册表编辑器打开

        # 模拟打开文件菜单的按键操作
        pyautogui.hotkey('alt', 'e')  # 打开文件菜单
        time.sleep(0.5)  # 等待菜单打开

        # 模拟按下 P 键以打开权限设置
        pyautogui.press('p')

    def on_tab_change(self, index):
        """当标签页改变时，刷新显示对应的图标"""
        if index == 0:
            self.selected_location = '此电脑'
        elif index == 1:
            self.selected_location = '资源管理器侧边栏'
        self.display_icons()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DriveIconManager()
    ex.show()
    sys.exit(app.exec())
