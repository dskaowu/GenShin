import os
import win32com.client
from configobj import ConfigObj


def check_exsit(process_name):  # 检查进程是否存在
    WMI = win32com.client.GetObject('winmgmts:')
    processCodeCov = WMI.ExecQuery(
        'select * from Win32_Process where Name like "%{}%"'.format(process_name))
    if len(processCodeCov) > 0:
        return True
    else:
        return False


def check():
    check1 = check_exsit('launcher.exe')
    check2 = check_exsit('YuanShen.exe')
    while (check1 or check2):
        print('请关闭原神启动器/客户端（关闭后按回车键继续）')
        input()
        check1 = check_exsit('launcher.exe')
        check2 = check_exsit('YuanShen.exe')


def get_now():  # 确认当前客户端服务器状态
    filename = './Genshin Impact/config.ini'
    config = ConfigObj(filename, encoding='UTF8')
    # print(config)
    if config['General']['cps'] == ('mihoyo' or 'pcadbdpz'):
        return '官服'
    if config['General']['cps'] == 'bilibili':
        return 'b服'


path = os.getcwd()


def copy(form_file, to_file):
    os.system('copy "' + path + form_file + '" "' + path + to_file + '" /y')


def xcopy(form_file, to_file):
    os.system('xcopy "' + path + form_file + '" "' + path + to_file + '"/E /y')


files_now = [r'\Genshin Impact\config.ini',
             r'\Genshin Impact\launcher.exe',
             r'\Genshin Impact\Genshin Impact Game\config.ini',
             r'\Genshin Impact\Genshin Impact Game\YuanShen_Data\Managed\Metadata\global-metadata.dat',
             r'\Genshin Impact\Genshin Impact Game\YuanShen_Data\Native\UserAssembly.dll',
             r'\Genshin Impact\Genshin Impact Game\YuanShen_Data\Native\UserAssembly.exp']
files_gf = [r'\exchange\gf\Genshin Impact',
            r'\exchange\gf\Genshin Impact',
            r'\exchange\gf\Genshin Impact\Genshin Impact Game',
            r'\exchange\gf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Managed\Metadata',
            r'\exchange\gf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Native',
            r'\exchange\gf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Native']
files_bf = [r'\exchange\bf\Genshin Impact',
            r'\exchange\bf\Genshin Impact',
            r'\exchange\bf\Genshin Impact\Genshin Impact Game',
            r'\exchange\bf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Managed\Metadata',
            r'\exchange\bf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Native',
            r'\exchange\bf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Native']
folder_now = r'\Genshin Impact\Genshin Impact Game\YuanShen_Data\Persistent'
folder_gf = r'\exchange\gf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Persistent'
folder_bf = r'\exchange\bf\Genshin Impact\Genshin Impact Game\YuanShen_Data\Persistent'


def backup(server_now):  # 备份当前客户端服务器状态
    if server_now == '官服':
        for i in range(6):
            copy(files_now[i], files_gf[i])
        xcopy(folder_now, folder_gf)
    if server_now == 'b服':
        for i in range(6):
            copy(files_now[i], files_bf[i])
        xcopy(folder_now, folder_bf)


def change(server_now):  # 更改客户端文件
    if server_now == '官服':
        os.system(
            r'xcopy ".\exchange\bf\Genshin Impact" ".\Genshin Impact" /E /y')
    if server_now == 'b服':
        os.system(
            r'xcopy ".\exchange\gf\Genshin Impact" ".\Genshin Impact" /E /y')


def if_first_boot():  # 检查是否首次启动
    if os.path.isfile('./exchange/first_boot'):
        check()
        print('第一次启动，自动更改游戏备份文件路径')
        config_now = ConfigObj('./Genshin Impact/config.ini', encoding='UTF8')
        game_path = config_now['General']['game_install_path']
        config_gf = ConfigObj(
            r'.\exchange\gf\Genshin Impact\config.ini', encoding='UTF8')
        # print(config_gf)
        config_gf['General']['game_install_path'] = game_path
        config_gf.write()
        # print(config_gf)
        config_bf = ConfigObj(
            r'.\exchange\bf\Genshin Impact\config.ini', encoding='UTF8')
        # print(config_bf)
        config_bf['General']['game_install_path'] = game_path
        config_bf.write()
        # print(config_bf)
        print('更改游戏备份文件路径完成')
        os.remove('./exchange/first_boot')


if __name__ == '__main__':
    check()
    if_first_boot()
    server_now = get_now()
    print('现在的服务器是{}'.format(server_now))
    backup(server_now)
    print('{}文件备份完毕'.format(server_now))
    change(server_now)
    server_now = get_now()
    print('已切换到{}'.format(server_now))
    input('输入回车退出')
