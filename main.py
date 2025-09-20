
import sys
import os
import shutil
import time
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import Qt
import qfluentwidgets as qfw
import psutil
from datetime import datetime
from PyQt6.QtWidgets import QFileDialog, QMessageBox
import ctypes


def show_error_and_wait(message):
    '''显示错误信息并等待用户查看'''
    print(f"[错误] {message}")
    print("程序将在10秒后自动关闭...")
    time.sleep(10)
    sys.exit(1)


class EtsToolInstaller(qfw.FluentWindow):
    def __init__(self):
        super().__init__()
        try:
            # 请求管理员权限
            self.request_admin_privileges()
            self.initUI()
            self.set_window_icon()
        except Exception as e:
            self.log(f"初始化失败: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            show_error_and_wait(f"程序初始化失败: {str(e)}")

    def request_admin_privileges(self):
        '''请求管理员权限'''
        try:
            if sys.platform == "win32":
                if not ctypes.windll.shell32.IsUserAnAdmin():
                    # 仅在非调试模式下请求权限提升
                    if not hasattr(sys, 'gettrace') or sys.gettrace() is None:
                        self.log("请求管理员权限...")
                        ctypes.windll.shell32.ShellExecuteW(
                            None, 
                            "runas", 
                            sys.executable, 
                            " ".join(sys.argv), 
                            None, 
                            1
                        )
                        sys.exit()
        except Exception as e:
            self.log(f"请求管理员权限时出错: {str(e)}")
            # 即使权限请求失败也继续运行程序

    def initUI(self):
        try:
            qfw.setTheme(qfw.Theme.DARK)
            self.setWindowTitle('E听说外挂工具安装器')
            self.setGeometry(100, 100, 800, 600)
            
            # 创建主界面部件
            self.main_widget = QWidget()
            self.main_widget.setObjectName("mainInterface")  # 设置对象名称
            self.main_layout = QVBoxLayout(self.main_widget)
            self.main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
            # 添加标题
            title_label = qfw.TitleLabel('E听说外挂工具安装器')
            title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.main_layout.addWidget(title_label)
            
            # 添加按钮区域
            button_layout = QHBoxLayout()
            button_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
            # 安装按钮
            install_button = qfw.PrimaryPushButton('安装', self.main_widget)
            install_button.setFixedWidth(200)
            install_button.clicked.connect(self.install)
            button_layout.addWidget(install_button)
            
            # 卸载按钮
            uninstall_button = qfw.PrimaryPushButton('卸载', self.main_widget)
            uninstall_button.setFixedWidth(200)
            uninstall_button.clicked.connect(self.uninstall)
            button_layout.addWidget(uninstall_button)
            
            self.main_layout.addLayout(button_layout)
            
            # 添加到导航界面
            self.addSubInterface(self.main_widget, qfw.FluentIcon.HOME, '开始')
            
            # 添加关于界面
            self.about_widget = QWidget()
            self.about_widget.setObjectName("aboutInterface")  # 设置对象名称
            about_layout = QVBoxLayout(self.about_widget)
            about_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
            about_label = qfw.SubtitleLabel('关于 E听说外挂工具')
            about_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            about_layout.addWidget(about_label)
            
            description = qfw.BodyLabel('这是一个用于安装和卸载 E听说外挂工具的程序。')
            description.setAlignment(Qt.AlignmentFlag.AlignCenter)
            about_layout.addWidget(description)
            
            self.addSubInterface(self.about_widget, qfw.FluentIcon.INFO, '关于')
        except Exception as e:
            self.log(f"初始化UI失败: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            show_error_and_wait(f"初始化UI失败: {str(e)}")

    def set_window_icon(self):
        '''设置窗口图标 - 适配PyInstaller打包'''
        try:
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
                icon_path1 = os.path.join(exe_dir, 'icon.ico')
                temp_path = sys._MEIPASS
                icon_path2 = os.path.join(temp_path, 'icon.ico')
                if os.path.exists(icon_path1):
                    self.log(f'''成功找到图标文件: {icon_path1}''')
                    self.setWindowIcon(QIcon(icon_path1))
                    self.log('窗口图标设置成功')
                    return None
                if os.path.exists(icon_path2):
                    self.log(f'''成功找到图标文件: {icon_path2}''')
                    self.setWindowIcon(QIcon(icon_path2))
                    self.log('窗口图标设置成功')
                    return None
                self.log(f'''未找到图标文件: {icon_path1} 或 {icon_path2}''')
                return None
            base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, 'icon.ico')
            if os.path.exists(icon_path):
                self.log(f'''成功找到图标文件: {icon_path}''')
                self.setWindowIcon(QIcon(icon_path))
                self.log('窗口图标设置成功')
                return None
            self.log(f'''未找到图标文件: {icon_path}''')
            return None
        except Exception as e:
            self.log(f"设置窗口图标失败: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            return None

    def log(self, message):
        '''记录操作日志 - 日志文件固定放在C盘根目录'''
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_message = f'''[{timestamp}] {message}\n'''
            log_path = 'C:\\ets_tool_log.txt'
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(log_message)
            # 同时输出到控制台
            print(log_message.strip())
        except Exception as e:
            print(f"写入日志失败: {e}")

    def check_target_processes(self):
        '''检查Ets.exe或shell.exe进程是否在运行'''
        try:
            target_processes = [
                'Ets.exe',
                'shell.exe']
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] in target_processes:
                    return proc.info['pid']
            return None
        except Exception as e:
            self.log(f"检查目标进程失败: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            return None

    def get_process_path(self, pid):
        '''获取进程的路径，并提供更详细的错误信息'''
        try:
            if pid is None:
                return None
            proc = psutil.Process(pid)
            exe_path = proc.exe()
            self.log(f'''成功获取进程PID {pid}的路径: {exe_path}''')
            return os.path.dirname(exe_path)
        except Exception as e:
            self.log(f'''获取进程路径失败: {str(e)}''')
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            return None

    def search_ets_folder(self):
        '''搜索系统中名为ETS或类似名称的文件夹'''
        try:
            self.log('开始搜索ETS相关文件夹')
            common_paths = [
                os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), '*ets*'),
                os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), '*ets*'),
                os.path.join(os.environ.get('LOCALAPPDATA', 'C:\\Users\\%USERNAME%\\AppData\\Local'), '*ets*'),
                os.path.join(os.environ.get('APPDATA', 'C:\\Users\\%USERNAME%\\AppData\\Roaming'), '*ets*'),
                'C:\\*ets*']
            
            import glob
            for pattern in common_paths:
                # 替换环境变量占位符
                pattern = pattern.replace('%USERNAME%', os.environ.get('USERNAME', ''))
                matches = glob.glob(pattern)
                for match in matches:
                    if os.path.isdir(match) and 'ets' in match.lower():
                        self.log(f'''找到ETS相关目录: {match}''')
                        return match
            return None
        except Exception as e:
            self.log(f"搜索ETS文件夹失败: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            return None

    def select_program_path(self):
        '''让用户手动选择Ets.exe或shell.exe文件'''
        try:
            self.log('启动手动选择程序路径功能')
            file_path, _ = QFileDialog.getOpenFileName(
                self, 
                '选择Ets.exe或shell.exe文件', 
                '', 
                '可执行文件 (*.exe);;所有文件 (*.*)'
            )
            if file_path:
                if 'Ets.exe' in file_path or 'shell.exe' in file_path:
                    proc_dir = os.path.dirname(file_path)
                    self.log(f'''用户手动选择的程序路径: {proc_dir}''')
                    return proc_dir
                else:
                    QMessageBox.warning(self, '警告', '请选择Ets.exe或shell.exe文件！')
                    self.log('用户选择了无效的程序文件')
                    return self.select_program_path()
            else:
                self.log('用户取消了程序路径选择')
                return None
        except Exception as e:
            self.log(f"手动选择程序路径失败: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            return None

    def install(self):
        '''安装E听说辅助工具 - 强制自动检测'''
        try:
            self.log('开始安装')
            proc_dir = None
            pid = self.check_target_processes()
            if pid:
                proc_dir = self.get_process_path(pid)
            if not proc_dir:
                self.log('程序未运行，正在搜索安装位置...')
                proc_dir = self.search_ets_folder()
            if not proc_dir:
                self.log('无法检测到E听说程序，请求用户手动选择')
                proc_dir = self.select_program_path()
            if not proc_dir:
                QMessageBox.critical(self, '错误', '无法自动检测到E听说程序的安装位置。请确保E听说已正确安装，然后重试。')
                self.log('安装失败：无法自动检测到程序路径')
                return None
            self.log(f'''最终使用的程序路径: {proc_dir}''')
            resource_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Resource')
            etstoolbox_src = os.path.join(resource_dir, 'etstoolbox')
            etstoolbox_dst = os.path.join(proc_dir, 'etstoolbox')
            winmm_src = os.path.join(resource_dir, 'winmm.dll')
            winmm_dst = os.path.join(proc_dir, 'winmm.dll')
            syswow64_dir = 'C:\\Windows\\SysWOW64'
            system32_dir = 'C:\\Windows\\System32'
            if os.path.exists(etstoolbox_dst):
                shutil.rmtree(etstoolbox_dst)
            shutil.copytree(etstoolbox_src, etstoolbox_dst)
            self.log(f'''复制etstoolbox文件夹到: {etstoolbox_dst}''')
            shutil.copy2(winmm_src, winmm_dst)
            self.log(f'''复制winmm.dll到: {winmm_dst}''')
            dll_files = [
                'msvcp140d.dll',
                'ucrtbased.dll',
                'vcruntime140d.dll']
            dll_success = True
            for dll_file in dll_files:
                try:
                    syswow64_path = os.path.join(syswow64_dir, dll_file)
                    system32_path = os.path.join(system32_dir, dll_file)
                    dll_src = os.path.join(resource_dir, dll_file)
                    if os.path.exists(dll_src):
                        if os.path.exists(syswow64_dir):
                            shutil.copy2(dll_src, syswow64_path)
                            self.log(f'''复制{dll_file}到: {syswow64_path}''')
                        if os.path.exists(system32_dir):
                            shutil.copy2(dll_src, system32_path)
                            self.log(f'''复制{dll_file}到: {system32_path}''')
                except Exception as e:
                    self.log(f'''复制{dll_file}失败: {str(e)}''')
                    import traceback
                    self.log(f"详细错误信息: {traceback.format_exc()}")
                    dll_success = False
            if dll_success:
                QMessageBox.information(self, '成功', '安装完成！')
                self.log('安装成功完成')
            else:
                QMessageBox.warning(self, '警告', '安装完成，但部分DLL文件复制失败！')
                self.log('安装完成，但部分DLL文件复制失败')
        except Exception as e:
            self.log(f"安装过程出现异常: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            QMessageBox.critical(self, '错误', f'安装过程中出现异常：{str(e)}')

    def uninstall(self):
        '''卸载E听说辅助工具 - 强制自动检测，仅在检测到目录时才允许删除系统DLL'''
        try:
            self.log('开始卸载')
            proc_dir = None
            pid = self.check_target_processes()
            if pid:
                proc_dir = self.get_process_path(pid)
            if not proc_dir:
                self.log('程序未运行，正在搜索安装位置...')
                proc_dir = self.search_ets_folder()
            if not proc_dir:
                self.log('无法检测到E听说程序，请求用户手动选择')
                proc_dir = self.select_program_path()
            if not proc_dir:
                QMessageBox.critical(self, '错误', '无法自动检测到E听说程序的安装位置。请确保E听说已正确安装，然后重试。')
                self.log('卸载失败：无法自动检测到程序路径')
                return None
            self.log(f'''最终使用的程序路径: {proc_dir}''')
            syswow64_dir = 'C:\\Windows\\SysWOW64'
            system32_dir = 'C:\\Windows\\System32'
            etstoolbox_dst = os.path.join(proc_dir, 'etstoolbox')
            winmm_dst = os.path.join(proc_dir, 'winmm.dll')
            if os.path.exists(etstoolbox_dst):
                shutil.rmtree(etstoolbox_dst)
                self.log(f'''删除etstoolbox文件夹: {etstoolbox_dst}''')
            if os.path.exists(winmm_dst):
                os.remove(winmm_dst)
                self.log(f'''删除winmm.dll: {winmm_dst}''')
            dll_files = [
                'msvcp140d.dll',
                'ucrtbased.dll',
                'vcruntime140d.dll']
            dll_success = True
            for dll_file in dll_files:
                try:
                    syswow64_path = os.path.join(syswow64_dir, dll_file)
                    if os.path.exists(syswow64_path):
                        os.remove(syswow64_path)
                        self.log(f'''从SysWOW64删除{dll_file}''')
                    system32_path = os.path.join(system32_dir, dll_file)
                    if os.path.exists(system32_path):
                        os.remove(system32_path)
                        self.log(f'''从System32删除{dll_file}''')
                except Exception as e:
                    self.log(f'''删除{dll_file}失败: {str(e)}''')
                    import traceback
                    self.log(f"详细错误信息: {traceback.format_exc()}")
                    dll_success = False
            if dll_success:
                QMessageBox.information(self, '成功', '卸载完成！')
                self.log('卸载成功完成')
            else:
                QMessageBox.warning(self, '警告', '卸载完成，但部分DLL文件删除失败！')
                self.log('卸载完成，但部分DLL文件删除失败')
        except Exception as e:
            self.log(f"卸载过程出现异常: {str(e)}")
            import traceback
            self.log(f"详细错误信息: {traceback.format_exc()}")
            QMessageBox.critical(self, '错误', f'卸载过程中出现异常：{str(e)}')


if __name__ == '__main__':
    try:
        print("程序启动中...")
        app = QApplication(sys.argv)
        ex = EtsToolInstaller()
        ex.show()
        print("程序启动成功")
        sys.exit(app.exec())
    except Exception as e:
        error_msg = f"程序启动失败: {str(e)}"
        print(error_msg)
        import traceback
        traceback.print_exc()
        
        # 记录到日志
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_message = f'''[{timestamp}] {error_msg}\n'''
            log_path = 'C:\\ets_tool_log.txt'
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(log_message)
                f.write(traceback.format_exc())
        except:
            pass
        
        # 显示错误并等待
        show_error_and_wait(error_msg)