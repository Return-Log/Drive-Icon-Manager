"""
https://github.com/Return-Log/Drive-Icon-Manager
GPL-3.0 license
coding: UTF-8
"""

import winreg
import ctypes
import win32api
import sys
import subprocess
import pyperclip
import win32security
import win32con
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QMessageBox, \
    QTabWidget, QListWidget, QListWidgetItem, QTextEdit
from PyQt6.QtCore import Qt
from RegistryPermissionsManager import RegistryPermissionsManager  # 修改注册表权限的模块

# 定义版本号和链接
VERSION = "v2.0"
GITHUB_LINK = "https://github.com/Return-Log/Drive-Icon-Manager"
FORUM_LINK = "https://www.52pojie.cn/thread-1955346-1-1.html"


class OutputRedirect:
    def __init__(self, text_edit):
        self.text_edit = text_edit

    def write(self, message):
        self.text_edit.append(message)  # 将消息添加到 QTextEdit 控件中

    def flush(self):
        pass

class DriveIconManager(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.icons = []
        self.selected_location = '此电脑'  # 默认选择“此电脑”
        self.display_icons()

        # 将标准输出和标准错误重定向到 GUI 控制台
        sys.stdout = OutputRedirect(self.terminal_output)
        sys.stderr = OutputRedirect(self.terminal_output)

    def initUI(self):
        self.setWindowTitle('Drive Icon Manager')
        self.setGeometry(150, 150, 500, 500)  # 增加高度以容纳新的控件

        layout = QVBoxLayout()

        # 欢迎和链接标签
        self.welcome_label = QLabel(f"欢迎使用 Drive Icon Manager 版本 {VERSION}", self)
        layout.addWidget(self.welcome_label)

        # 链接布局
        link_layout = QHBoxLayout()
        self.github_link = QLabel(f"<a href='{GITHUB_LINK}'>访问 GitHub 仓库</a>", self)
        self.github_link.setOpenExternalLinks(True)
        link_layout.addWidget(self.github_link)

        self.forum_link = QLabel(f"<a href='{FORUM_LINK}'>访问 吾爱破解论坛 寻求帮助</a>", self)
        self.forum_link.setOpenExternalLinks(True)
        link_layout.addWidget(self.forum_link)

        layout.addLayout(link_layout)

        # 不同位置的标签页
        self.tab_widget = QTabWidget(self)
        self.this_pc_tab = QWidget()
        self.sidebar_tab = QWidget()
        self.permissions_tab = QWidget()

        self.tab_widget.addTab(self.this_pc_tab, "此电脑")
        self.tab_widget.addTab(self.sidebar_tab, "资源管理器侧边栏")
        self.tab_widget.addTab(self.permissions_tab, "注册表权限")
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

        self.permissions_layout = QVBoxLayout()
        self.permissions_list = QListWidget()
        self.permissions_layout.addWidget(self.permissions_list)

        # 添加按钮和终端输出区域
        self.open_this_pc_button = QPushButton("打开此电脑注册表", self)
        self.open_this_pc_button.clicked.connect(self.open_this_pc_registry)
        self.permissions_layout.addWidget(self.open_this_pc_button)

        self.open_sidebar_button = QPushButton("打开资源管理器侧边栏注册表", self)
        self.open_sidebar_button.clicked.connect(self.open_sidebar_registry)
        self.permissions_layout.addWidget(self.open_sidebar_button)

        self.terminal_output = QTextEdit()
        self.terminal_output.setReadOnly(True)
        self.permissions_layout.addWidget(self.terminal_output)

        self.permissions_tab.setLayout(self.permissions_layout)

        # 向布局添加标签
        layout.addWidget(self.tab_widget)

        # 显示权限状态
        self.display_permissions()

        # 刷新和删除按钮的水平布局
        button_layout = QHBoxLayout()

        self.refresh_button = QPushButton('刷新', self)
        self.refresh_button.clicked.connect(self.display_icons)
        self.refresh_button.setFixedWidth(60)
        button_layout.addWidget(self.refresh_button)

        self.delete_button = QPushButton('删除选中的驱动器图标', self)
        self.delete_button.clicked.connect(self.delete_selected_icon)
        button_layout.addWidget(self.delete_button)

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
            base_key = winreg.HKEY_CURRENT_USER
            key_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
        else:
            base_key = winreg.HKEY_USERS
            current_user_sid = self.get_current_user_sid()
            key_path = fr'{current_user_sid}\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'

        try:
            key = winreg.OpenKey(base_key, key_path, 0, winreg.KEY_ALL_ACCESS)
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

    def get_current_user_name(self):
        """获取当前用户的用户名"""
        try:
            return win32api.GetUserName()
        except Exception as e:
            self.display_error_message(f"无法获取当前用户的用户名: {e}")
            return None

    def display_permissions(self):
        """显示 '此电脑' 和 '资源管理器侧边栏' 对应的注册表项的权限状态"""
        self.permissions_list.clear()

        # 定义路径和用户组
        this_pc_key_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
        sidebar_key_path = fr'{self.get_current_user_sid()}\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'
        user_group = self.get_current_user_name()
        access_rights = win32con.KEY_ALL_ACCESS | win32con.KEY_READ

        try:
            # 创建RegistryPermissionsManager实例，传入正确的参数
            this_pc_manager = RegistryPermissionsManager(
                root_key=win32con.HKEY_CURRENT_USER,
                key_path=this_pc_key_path,
                user_name=user_group,
                access_rights=access_rights
            )

            sidebar_manager = RegistryPermissionsManager(
                root_key=win32con.HKEY_USERS,
                key_path=sidebar_key_path,
                user_name=user_group,
                access_rights=access_rights
            )

            # 检查权限状态
            this_pc_permission = this_pc_manager.check_permissions()
            sidebar_permission = sidebar_manager.check_permissions()

            # 显示权限状态
            this_pc_item = QListWidgetItem(
                f"此电脑权限: {'写入权限已禁用' if this_pc_permission else '已启用写入权限'}")
            this_pc_item.setFlags(this_pc_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)  # 设置不可选择

            sidebar_item = QListWidgetItem(
                f"资源管理器侧边栏权限: {'写入权限已禁用' if sidebar_permission else '已启用写入权限'}")
            sidebar_item.setFlags(sidebar_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)  # 设置不可选择

            # 添加启用/禁用写入权限按钮
            this_pc_button = QPushButton("启用写入权限" if this_pc_permission else "禁用写入权限")
            this_pc_button.setFixedHeight(20)  # 设置按钮高度
            this_pc_button.clicked.connect(lambda: self.toggle_permission(this_pc_manager, this_pc_permission))
            self.permissions_list.addItem(this_pc_item)
            self.permissions_list.addItem(QListWidgetItem())  # 空项用于分隔按钮和权限状态
            self.permissions_list.setItemWidget(self.permissions_list.item(self.permissions_list.count() - 1),
                                                this_pc_button)

            sidebar_button = QPushButton("启用写入权限" if sidebar_permission else "禁用写入权限")
            sidebar_button.setFixedHeight(20)  # 设置按钮高度
            sidebar_button.clicked.connect(lambda: self.toggle_permission(sidebar_manager, sidebar_permission))
            self.permissions_list.addItem(sidebar_item)
            self.permissions_list.addItem(QListWidgetItem())  # 空项用于分隔按钮和权限状态
            self.permissions_list.setItemWidget(self.permissions_list.item(self.permissions_list.count() - 1),
                                                sidebar_button)

        except Exception as e:
            print(f"发生错误: {e}")
            error_item = QListWidgetItem(f"无法读取注册表项权限。错误信息: {e}")
            error_item.setFlags(error_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)  # 设置不可选择
            self.permissions_list.addItem(error_item)

    def toggle_permission(self, manager, current_permission):
        """切换权限状态"""
        if current_permission:
            # 当前禁用了写入权限，恢复权限
            manager.modify_permissions(deny=False)
        else:
            # 当前启用了写入权限，禁用权限
            manager.modify_permissions(deny=True)

        # 操作完成后刷新权限显示
        self.display_permissions()

    def open_this_pc_registry(self):
        """打开此电脑注册表路径并复制路径"""
        path = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"
        self.open_registry_editor(path)

    def open_sidebar_registry(self):
        """打开资源管理器侧边栏注册表路径并复制路径"""
        sid = self.get_current_user_sid()
        path = fr'HKEY_USERS\{sid}\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'
        self.open_registry_editor(path)

    def open_registry_editor(self, path):
        """打开注册表编辑器并复制路径"""
        try:
            # 复制路径到剪贴板
            pyperclip.copy(path)
            self.log_terminal_output(f"已复制路径到剪贴板: {path}")

            # 打开注册表编辑器
            subprocess.run(["regedit.exe"])
        except Exception as e:
            self.log_terminal_output(f"无法打开注册表编辑器: {e}")

    def log_terminal_output(self, message):
        """将消息输出到终端区域"""
        self.terminal_output.append(message)

    def display_error_message(self, message):
        """在当前标签页显示错误信息"""
        item = QListWidgetItem(message)
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsSelectable)  # 设置不可选择
        if self.selected_location == '此电脑':
            self.this_pc_text.addItem(item)
        elif self.selected_location == '资源管理器侧边栏':
            self.sidebar_text.addItem(item)

    def on_tab_change(self, index):
        """当标签页改变时，刷新显示对应的图标"""
        if index == 0:
            self.selected_location = '此电脑'
            self.display_icons()
        elif index == 1:
            self.selected_location = '资源管理器侧边栏'
            self.display_icons()
        elif index == 2:
            self.display_permissions()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DriveIconManager()
    ex.show()
    sys.exit(app.exec())
