# 网盘图标删除器 Drive Icon Manager

## 介绍

- 一键轻松删除Windows平台“此电脑”及“资源管理器侧边栏”中的第三方图标，使你不再受无用的网盘图标影响。

- 这些图标当然可以手动删除，但手动查找注册表是一件繁琐的事，此外“资源管理器侧边栏”中的图标路径包含用户独有的SID，直接复制网上的路径是行不通的，本程序会自动获取当前用户SID，节省自己查找的时间。

## 原理

使用管理员权限运行软件会自动读取注册表

- 【“此电脑”中的图标地址】HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\  

- 【“资源管理器侧边栏”中的图标地址】HKEY_USERS\S-1-5-21-652943116-554383430-1123514156-1000（此处SID每个用户都不同）\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace

并列出子项，手动选择需要删除的项删除

## 界面展示

![Snipaste_2024-08-16_22-38-22.png](https://s2.loli.net/2024/08/16/RsFi5n7M2PzZNdo.png)

## 开发环境

必要的库

```shell
pip install colorama
```

```shell
pip install pywin32
```

python版本：3.12

------

Copyright © 2024 Log All rights reserved.