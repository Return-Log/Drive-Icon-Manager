"""
https://github.com/Return-Log/Drive-Icon-Manager
GPL-3.0 license
coding: UTF-8
"""

import win32api
import win32security
import win32con


class RegistryPermissionsManager:
    def __init__(self, root_key, key_path, user_name, access_rights):
        self.root_key = root_key  # 根键 (如 HKEY_USERS, HKEY_LOCAL_MACHINE)
        self.key_path = key_path  # 注册表路径
        self.user_name = user_name  # 用户或组名
        self.access_rights = access_rights

    def modify_permissions(self, deny=True):
        """拒绝或恢复相应的注册表权限"""
        try:
            # 打开指定路径的注册表项
            key = win32api.RegOpenKeyEx(
                self.root_key,  # 可指定不同的根键
                self.key_path,  # 注册表项路径
                0,  # 保留参数，一般设为0
                win32con.KEY_ALL_ACCESS  # 允许对注册表项进行所有访问的权限
            )

            # 获取注册表项的安全描述符
            security_descriptor = win32security.GetSecurityInfo(
                key,  # 注册表项句柄
                win32security.SE_REGISTRY_KEY,  # 表示我们要操作的是注册表项
                win32security.DACL_SECURITY_INFORMATION  # 指定我们要获取DACL（自定义访问控制列表）
            )

            # 获取当前注册表项的 DACL（自由访问控制列表）
            dacl = security_descriptor.GetSecurityDescriptorDacl()

            # 获取指定用户的 SID（安全标识符）
            user_sid, _, _ = win32security.LookupAccountName(None, self.user_name)

            if deny:
                # 添加拒绝访问的 ACE（访问控制条目）到 DACL
                dacl.AddAccessDeniedAceEx(
                    win32security.ACL_REVISION,  # DACL 版本号
                    win32con.CONTAINER_INHERIT_ACE | win32con.OBJECT_INHERIT_ACE,  # 继承标志
                    self.access_rights,  # 访问权限（如完全控制和读取）
                    user_sid  # 目标用户的 SID
                )
                print(f"权限被: {self.user_name} 禁用成功，路径： {self.key_path}.")
            else:
                # 逆序遍历 DACL，删除与用户相关的所有拒绝 ACE
                ace_count = dacl.GetAceCount()  # 获取 DACL 中 ACE 的数量
                for ace_index in range(ace_count - 1, -1, -1):
                    ace = dacl.GetAce(ace_index)
                    ace_type = ace[0][0]
                    ace_sid = ace[2]

                    if ace_type == win32security.ACCESS_DENIED_ACE_TYPE and ace_sid == user_sid:
                        dacl.DeleteAce(ace_index)  # 删除指定的 ACE
                        print(f"已删除被拒绝的 ACE: {self.user_name} 在 {self.key_path}.")

                print(f"权限已被: {self.user_name} 恢复，路径： {self.key_path}.")

            # 将修改后的 DACL 设置回注册表项的安全描述符
            win32security.SetSecurityInfo(
                key,  # 注册表项句柄
                win32security.SE_REGISTRY_KEY,  # 操作对象类型为注册表项
                win32security.DACL_SECURITY_INFORMATION,  # 表示我们要设置 DACL
                None,  # 不修改所有者
                None,  # 不修改组
                dacl,  # 更新的 DACL
                None  # 不修改 SACL（系统访问控制列表）
            )

            # 关闭注册表项句柄
            win32api.RegCloseKey(key)

        except Exception as e:
            print(f"发生错误: {e}")

    def check_permissions(self):
        """检查注册表权限状态"""
        try:
            # 打开指定路径的注册表项
            key = win32api.RegOpenKeyEx(
                self.root_key,  # 可指定不同的根键
                self.key_path,  # 注册表项路径
                0,  # 保留参数，一般设为0
                win32con.KEY_READ  # 只读权限打开注册表项
            )

            # 获取注册表项的安全描述符
            security_descriptor = win32security.GetSecurityInfo(
                key,  # 注册表项句柄
                win32security.SE_REGISTRY_KEY,  # 表示我们要操作的是注册表项
                win32security.DACL_SECURITY_INFORMATION  # 指定我们要获取DACL
            )

            # 获取当前注册表项的 DACL
            dacl = security_descriptor.GetSecurityDescriptorDacl()

            # 获取指定用户的 SID
            user_sid, _, _ = win32security.LookupAccountName(None, self.user_name)

            has_deny = False
            has_allow = False

            # 遍历 DACL 中的 ACE，检查是否有拒绝或允许的条目
            for ace_index in range(dacl.GetAceCount()):
                ace = dacl.GetAce(ace_index)
                ace_type = ace[0][0]
                ace_mask = ace[1]
                ace_sid = ace[2]

                if ace_sid == user_sid:
                    if ace_type == win32security.ACCESS_DENIED_ACE_TYPE and ace_mask & self.access_rights == self.access_rights:
                        has_deny = True
                    if ace_type == win32security.ACCESS_ALLOWED_ACE_TYPE and ace_mask & self.access_rights == self.access_rights:
                        has_allow = True

            # 关闭注册表项句柄
            win32api.RegCloseKey(key)

            if has_deny and not has_allow:
                return True  # 已拒绝完全控制和读取权限
            elif has_allow and not has_deny:
                return False  # 已恢复完全控制和读取权限
            else:
                return True  # 状态不明确或未设置

        except Exception as e:
            print(f"注册表权限已被禁用，如需恢复点击打开注册表，【路径自动复制】在地址栏粘贴路径后将NameSpace文件夹当前用户权限中的‘拒绝’取消勾选。发生错误: {e}")
            return None


# 示例用法
if __name__ == "__main__":
    registry_root_key = win32con.HKEY_CURRENT_USER  # 可以修改为 HKEY_USERS, HKEY_LOCAL_MACHINE 等
    registry_key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"  # 注册表路径
    user_name = "SYSTEM"  # 用户或组名
    access_rights = win32con.KEY_ALL_ACCESS | win32con.KEY_READ  # 默认权限

    # 创建注册表权限管理器实例
    rpm = RegistryPermissionsManager(registry_root_key, registry_key_path, user_name, access_rights)

    # # 拒绝完全控制和读取权限
    # rpm.modify_permissions(deny=True)
    #
    # # 检查是否已拒绝完全控制和读取权限
    # is_denied = rpm.check_permissions()
    # print(f"权限被拒绝: {is_denied}")
    #
    # # 恢复完全控制和读取权限
    # rpm.modify_permissions(deny=False)
    #
    # # 检查是否已恢复完全控制和读取权限
    # is_restored = not rpm.check_permissions()
    # print(f"权限已恢复: {is_restored}")
