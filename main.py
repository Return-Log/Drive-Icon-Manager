"""
https://github.com/Return-Log/Drive-Icon-Manager
GPL-3.0 license
coding: UTF-8
"""
import os
import winreg
import ctypes
import win32api
import sys
import subprocess
import pyperclip
import win32security
import win32con
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QMessageBox, \
    QTabWidget, QListWidget, QListWidgetItem, QTextEdit, QTextBrowser
from PyQt6.QtCore import Qt
from markdown import markdown
from RegistryPermissionsManager import RegistryPermissionsManager  # 修改注册表权限的模块

# 定义版本号和链接
VERSION = "v2.2"
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
        self.setGeometry(150, 150, 500, 500)

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
        self.about_tab = QWidget()

        self.tab_widget.addTab(self.this_pc_tab, "此电脑")
        self.tab_widget.addTab(self.sidebar_tab, "资源管理器侧边栏")
        self.tab_widget.addTab(self.permissions_tab, "注册表权限")
        self.tab_widget.addTab(self.about_tab, "关于")
        self.tab_widget.currentChanged.connect(self.on_tab_change)

        # 此电脑标签页的布局
        self.this_pc_layout = QVBoxLayout()
        self.this_pc_text = QListWidget()
        self.this_pc_layout.addWidget(self.this_pc_text)
        self.this_pc_tab.setLayout(self.this_pc_layout)

        # 资源管理器侧边栏标签页的布局
        self.sidebar_layout = QVBoxLayout()
        self.sidebar_text = QListWidget()
        self.sidebar_layout.addWidget(self.sidebar_text)
        self.sidebar_tab.setLayout(self.sidebar_layout)

        # 注册表权限标签页的布局
        self.permissions_layout = QVBoxLayout()
        self.permissions_list = QListWidget()
        self.permissions_layout.addWidget(self.permissions_list)

        # 关于标签页的布局
        self.about_layout = QVBoxLayout()
        self.about_text_browser = QTextBrowser()

        # ‘关于’界面的按钮
        self.open_this_pc_button = QPushButton("打开此电脑注册表", self)
        self.open_this_pc_button.clicked.connect(self.open_this_pc_registry)
        self.permissions_layout.addWidget(self.open_this_pc_button)

        self.open_this_pc_button_2 = QPushButton("打开此电脑2注册表", self)
        self.open_this_pc_button_2.clicked.connect(self.open_this_pc_registry_2)
        self.permissions_layout.addWidget(self.open_this_pc_button_2)

        self.open_sidebar_button = QPushButton("打开资源管理器侧边栏注册表", self)
        self.open_sidebar_button.clicked.connect(self.open_sidebar_registry)
        self.permissions_layout.addWidget(self.open_sidebar_button)

        # 终端输出
        self.terminal_output = QTextEdit()
        self.terminal_output.setReadOnly(True)
        self.permissions_layout.addWidget(self.terminal_output)

        self.permissions_tab.setLayout(self.permissions_layout)

        # 向布局添加标签
        layout.addWidget(self.tab_widget)

        # 显示权限状态
        self.display_permissions()

        # 加载并显示关于内容
        about_content = self.load_about_content()
        try:
            html_content = markdown(about_content)
            self.about_text_browser.setHtml(html_content)
        except Exception as e:
            self.about_text_browser.setText(f"无法处理关于内容: {e}")

        self.about_layout.addWidget(self.about_text_browser)
        self.about_tab.setLayout(self.about_layout)

        # 底部按钮的水平布局
        button_layout = QHBoxLayout()

        self.refresh_button = QPushButton('刷新', self)
        self.refresh_button.clicked.connect(self.display_icons)
        self.refresh_button.setFixedWidth(60)
        button_layout.addWidget(self.refresh_button)

        self.delete_button = QPushButton('删除选中的驱动器图标', self)
        self.delete_button.clicked.connect(self.delete_selected_icon)
        button_layout.addWidget(self.delete_button)

        # 添加备份按钮
        self.backup_button = QPushButton('备份选中的驱动器图标', self)
        self.backup_button.clicked.connect(self.backup_selected_icon)
        button_layout.addWidget(self.backup_button)

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

                    # 仅对 'HKEY_LOCAL_MACHINE' 进行过滤，显示符合条件的第三方程序图标
                    if base_key == winreg.HKEY_LOCAL_MACHINE and \
                            (display_name != "CLSID" in display_name or subkey_name == "DelegateFolders"):
                        winreg.CloseKey(subkey)
                        i += 1
                        continue

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
        if index < 0 or index >= len(self.icons):
            QMessageBox.warning(self, "未选中任何项", "请先选择一个驱动器图标再进行删除操作")
            return

        subkey_name, _, source = self.icons[index]
        if source == '此电脑':
            base_key = winreg.HKEY_CURRENT_USER
            key_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
        elif source == '此电脑2':
            base_key = winreg.HKEY_LOCAL_MACHINE
            key_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
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

    def backup_selected_icon(self):
        """备份选中的驱动器图标"""
        selected_index = self.this_pc_text.currentRow() if self.selected_location == '此电脑' else self.sidebar_text.currentRow()

        if selected_index < 0 or selected_index >= len(self.icons):
            QMessageBox.warning(self, "无效的选择", "请先选择一个有效的驱动器图标再进行操作")
            return

        subkey_name, display_name, source = self.icons[selected_index]

        # 根据源选择正确的根键和路径
        if source == '此电脑':
            base_key = winreg.HKEY_CURRENT_USER
            key_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
            base_key_str = 'HKEY_CURRENT_USER'
        elif source == '此电脑2':
            base_key = winreg.HKEY_LOCAL_MACHINE
            key_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
            base_key_str = 'HKEY_LOCAL_MACHINE'
        elif source == '资源管理器侧边栏':
            base_key = winreg.HKEY_USERS
            current_user_sid = self.get_current_user_sid()
            key_path = fr'{current_user_sid}\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'
            base_key_str = 'HKEY_USERS'

        full_key_path = f"{key_path}\\{subkey_name}"

        try:
            reg_file_name = f"双击恢复{display_name}-[{source}]图标.reg"
            with open(reg_file_name, 'w', encoding='utf-16le') as reg_file:
                reg_file.write("\ufeff")  # 写入BOM以标识UTF-16LE编码
                reg_file.write("Windows Registry Editor Version 5.00\n\n")
                reg_file.write(f"[{base_key_str}\\{full_key_path}]\n")

                key = winreg.OpenKey(base_key, full_key_path, 0, winreg.KEY_READ)
                i = 0
                while True:
                    try:
                        value_name, value_data, value_type = winreg.EnumValue(key, i)
                        if value_type == winreg.REG_SZ:
                            reg_file.write(f"\"{value_name}\"=\"{value_data}\"\n")
                        elif value_type == winreg.REG_DWORD:
                            reg_file.write(f"\"{value_name}\"=dword:{value_data:08x}\n")
                        elif value_type == winreg.REG_BINARY:
                            reg_file.write(f"\"{value_name}\"=hex:{','.join([f'{b:02x}' for b in value_data])}\n")
                        elif value_type == winreg.REG_MULTI_SZ:
                            reg_file.write(
                                f"\"{value_name}\"=hex(7):{','.join([f'{ord(c):02x}' for c in '\\0'.join(value_data)])},00,00\n")
                        elif value_type == winreg.REG_EXPAND_SZ:
                            reg_file.write(
                                f"\"{value_name}\"=hex(2):{','.join([f'{ord(c):02x}' for c in value_data])},00,00\n")
                        i += 1
                    except OSError:
                        break
                winreg.CloseKey(key)

            QMessageBox.information(self, "备份成功", f"驱动器图标已成功备份到 {reg_file_name}")

        except Exception as e:
            self.display_error_message(f"无法备份驱动器图标: {e}")

    def display_icons(self):
        """显示所有驱动器图标"""
        self.icons = []
        if self.selected_location == '此电脑':
            self.this_pc_text.clear()

            # 列出 HKEY_CURRENT_USER 下的驱动器图标
            self.icons += self.list_drive_icons(winreg.HKEY_CURRENT_USER,
                                                r'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace',
                                                '此电脑')

            # 列出 HKEY_LOCAL_MACHINE 下的驱动器图标
            self.icons += self.list_drive_icons(winreg.HKEY_LOCAL_MACHINE,
                                                r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace',
                                                '此电脑2')

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

    def open_this_pc_registry_2(self):
        """打开此电脑2注册表路径并复制路径"""
        path = r"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"
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

    def load_about_content(self):
        """加载关于内容"""
        try:
            if getattr(sys, 'frozen', False):
                # 如果是打包后的程序
                about_file_path = os.path.join(sys._MEIPASS, 'about_content.md')
            else:
                # 如果是开发环境
                about_file_path = 'about_content.md'

            with open(about_file_path, 'r', encoding='utf-8') as file:
                about_text = file.read()

            return about_text if about_text else "关于内容为空。"

        except FileNotFoundError:
            return "关于内容文件未找到。"


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DriveIconManager()
    ex.show()
    sys.exit(app.exec())
