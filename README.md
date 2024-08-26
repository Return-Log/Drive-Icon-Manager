# 网盘图标删除器 Drive Icon Manager v2.2

## 介绍

- 一键轻松删除Windows平台“此电脑”及“资源管理器侧边栏”中的第三方图标，使你不再受无用的网盘图标影响。
- 这些图标当然可以手动删除，但手动查找注册表是一件繁琐的事，此外“资源管理器侧边栏”中的图标路径包含用户独有的SID，直接复制网上的路径是行不通的，本程序会自动获取当前用户SID，节省自己查找的时间。
- 一键锁定注册表，使图标不再复发
- 一键备份注册表项，有需要时可以恢复

## 原理

使用管理员权限运行软件会自动读取注册表

- 【“此电脑”中的图标地址】HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\  

- 【“资源管理器侧边栏”中的图标地址】HKEY_USERS\S-1-5-21-652943116-554383430-1123514156-1000（此处SID每个用户都不同）\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace

并列出子项，手动选择需要删除的项删除

## 界面展示

<img src="https://s2.loli.net/2024/08/24/LQqXGhWFZyK2Nzu.png" alt="Snipaste_2024-08-24_14-12-01.png" style="zoom: 67%;" />

<img src="https://s2.loli.net/2024/08/19/2uWPh1e9paBV5xQ.png" alt="Snipaste_2024-08-19_21-06-31.png" style="zoom: 67%;" />

<img src="https://s2.loli.net/2024/08/19/QojfFreUpyhEKuY.png" alt="Snipaste_2024-08-19_20-40-26.png" style="zoom: 67%;" />

## 详细说明

### 禁用写入权限

- 点击‘禁用写入权限’后其它软件无法读取和更改注册表，当然此程序也无法读取相应注册表，【务必】在禁用写入权限前先将无用的图标删除。

- 图标的删除是永久性的，请仔细确认防止误删。

- 点击了‘禁用写入权限’后如果想继续让其它软件添加图标只能进行手动操作，如下图所示

  ![Snipaste_2024-08-19_20-50-33.png](https://s2.loli.net/2024/08/19/Fr7NeGY6BwlDEqL.png)

### 备份注册表

- 选中需要备份的项，点击下方备份按钮，会在程序根目录下生成对应的.reg文件，需要恢复时，双击即可恢复

### 部分‘此电脑’中图标

- ‘此电脑’标签页显示两个一样的项时，需要都删除
- 另一个项可以点‘注册表权限’中的‘打开此电脑2注册表’快速打开

## 发行版

**运行环境：win10/win11 非精简版系统**

### 开源仓库

https://github.com/Return-Log/Drive-Icon-Manager

### 网盘下载

**v2.2 单文件版**【推荐下载】

https://wwif.lanzouk.com/iRLAj289qoqb
密码:heeq

================================

**v2.1 单文件版**

https://wwif.lanzouk.com/i4uja2862t4h
密码:e1oy

================================

**v2.0 单文件版**
https://wwif.lanzouh.com/iY3TQ27veitc 

密码:18kd

**v2.0 绿色版**
https://wwif.lanzouh.com/ivwHW27vel4f

=================================

**v1.1 单文件版**
https://wwif.lanzouh.com/iaj7727m1zqj 

密码:czzg

## 开发环境

必要的库

```python
import winreg
import ctypes
import win32api
import sys
import subprocess
import pyperclip
import win32security
import win32con
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QMessageBox, QTabWidget, QListWidget, QListWidgetItem, QTextEdit, QTextBrowser
from PyQt6.QtCore import Qt
from markdown import markdown
from RegistryPermissionsManager import RegistryPermissionsManager  # 修改注册表权限的模块
```

python版本：3.12

[![Star History Chart](https://api.star-history.com/svg?repos=Return-Log/Drive-Icon-Manager&type=Date)](https://star-history.com/#Return-Log/Drive-Icon-Manager&Date)

------

Copyright © 2024 Log All rights reserved.