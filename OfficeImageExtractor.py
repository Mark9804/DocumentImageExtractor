# coding:utf-8
import os
import sys
import zipfile
import shutil
import subprocess
from re import sub
from pymsgbox import alert
from locale import getdefaultlocale

UserLocale = 'undefined'
#########错误信息#########
errormsg_international = {
    'ErrorOnLoad': 'Drag & drop an Microsoft Office file on me:(',
    'ErrorOnLoadTitle': 'OfficeImageExtractor: Error Occurred',
    'ErrorOldOfficeFormat': 'Earlier Office format file is not supported yet:( Try saving as ',
    'ErrorOldOfficeFormatSuffix': ' file and then retry.',
    'ErrorNoMediaFile': 'This document does not seem to contain any media file.',
    'ErrorNoMediaFileTitle': 'OfficeImageExtractor: No Media Found',
    'ErrorOnRemoveTemp': 'The attempt to remove',
    'ErrorOnRemoveEmpty': 'The attempt to remove',
    'ErrorOnRemoveManual': 'failed. Please remove it manually.',
    'ErrorOnRemoveTitle': 'OfficeImageExtractor: Error Occurred During File Deletion',
    'ErrorOnRemoveEmptyTitle': 'OfficeImageExtractor: Error Occurred During File Deletion',
    'ErrorOnRemoveTip': 'DELETE_THIS_FOLDER',
}

errormsg_cn = {
    'ErrorOnLoad': '请将一个Office文档拖拽至程序上，然后松开鼠标左键',
    'ErrorOnLoadTitle': 'OfficeImageExtractor: 错误发生',
    'ErrorOldOfficeFormat': '暂时不支持较早格式的Office文档。请尝试转换为',
    'ErrorOldOfficeFormatSuffix': '格式后重试。',
    'ErrorNoMediaFile': '这个文档似乎不包含媒体文件。',
    'ErrorNoMediaFileTitle': 'OfficeImageExtractor: 未找到媒体文件',
    'ErrorOnRemoveTemp': '尝试清除缓存文件夹',
    'ErrorOnRemoveEmpty': '尝试清除空文件夹',
    'ErrorOnRemoveManual': '失败，请手动删除该文件夹',
    'ErrorOnRemoveTitle': 'OfficeImageExtractor: 删除临时文件时发生错误',
    'ErrorOnRemoveEmptyTitle': 'OfficeImageExtractor: 删除空文件夹时发生错误',
    'ErrorOnRemoveTip': '请手动删除该文件夹',
}  #########错误信息结束#########


def error(errorcode):
    global UserLocale
    if UserLocale == 'undefined':
        locale_lang, codepage = getdefaultlocale()
        if locale_lang == 'zh_CN':
            UserLocale = 'cn'
        else:
            UserLocale = 'international'
    else:
        pass
    globals()['errormsg'] = globals()['errormsg_' + str(UserLocale)]
    return errormsg[str(errorcode)]


def office2pic(OfficeDocPath, ZipFilePath, TempPath, StorePath, FileType):
    # 清理现场。NoMedia表示不含任何媒体文件，连最终文件夹一起删掉。
    def clean(NoMedia=False):
        try:
            subprocess.call('rmdir /s /q "' + TempPath + '"', shell=True)
            # os.remove(TempPath)  # Windows环境下会报权限错误，可能不能操作文件夹
        except FileNotFoundError:
            # 如果删除失败，就打开临时文件夹让用户自己删除
            alert(error('ErrorOnRemoveTemp') + ' "' + str(TempPath) + '" ' + error('ErrorOnRemoveManual'),
                  error('ErrorOnRemoveTitle'))
            with open(str(TempPath + '\\' + error('ErrorOnRemoveTip')), 'w'):
                pass
            subprocess.Popen('explorer "' + TempPath + '"', shell=False)
        if NoMedia:
            try:
                subprocess.call('rmdir /s /q "' + StorePath + '"', shell=True)
            except FileNotFoundError:
                alert(error('ErrorOnRemoveEmpty') + ' "' + str(StorePath) + '" ' + error('ErrorOnRemoveManual'),
                      error('ErrorOnRemoveEmptyTitle'))
                with open(str(TempPath + '\\' + error('ErrorOnRemoveTip')), 'w'):
                    pass
                subprocess.Popen('explorer "' + StorePath + '"', shell=False)

    shutil.copy(OfficeDocPath, ZipFilePath)  # 创建副本并修改后缀名为.zip
    file = zipfile.ZipFile(ZipFilePath, 'r')
    for files in file.namelist():
        file.extract(files, TempPath)  # 解压至临时文件夹
    file.close()  # 解除副本文件占用
    os.remove(ZipFilePath)  # 删除副本

    # 复制相关文件夹的图片至目标文件夹
    try:
        pictures = os.listdir(os.path.join(TempPath + '\\' + FileType + '\\media'))
    # 如果找不到media文件夹，说明没有媒体文件。删掉创建的文件夹之后报错退出
    except FileNotFoundError:
        alert(error('ErrorNoMediaFile'), error('ErrorNoMediaFileTitle'))
        clean(NoMedia=True)
        quit()
    # 191117: 这一步和下一步当中的rmdir不能正确处理带空格文件！！！ ——191203:虽然不知道为什么，但是现在可以了
    for picture in pictures:
        shutil.copy(os.path.join(TempPath + '\\' + FileType + '\\media\\', picture), StorePath)
    clean()
    subprocess.Popen('explorer "' + StorePath + '"', shell=False)


def makestorepath(InputFilePath):
    # 对输入路径进行处理，删除不合法字符
    zippath = sub(u'([^\u4e00-\u9fa5\u0030-\u0039\u0041-\u005a\u0061-\u007a])', '_', InputFilePath) + '.zip'
    storepath = sub(r'.(doc|ppt|xls)x$', '', InputFilePath)
    if not os.path.exists(storepath):
        os.mkdir(storepath)
    return zippath, storepath


if __name__ == '__main__':
    try:
        path = sys.argv[1]
    except:
        alert(error('ErrorOnLoad'), error('ErrorOnLoadTitle'))
        quit()
    temp_path = os.path.join(str(os.getcwd()) + '\\temp')
    # Done: 抽象化设置路径流程，使用re.sub解决strip误伤问题
    if str(path)[-5:] == '.docx':
        zip_filepath, store_path = makestorepath(sys.argv[1])
        office2pic(path, zip_filepath, temp_path, store_path, 'word')  # 最后一个参数是每个格式存放图片的目录

    elif str(path)[-5:] == '.pptx':
        zip_filepath, store_path = makestorepath(sys.argv[1])
        office2pic(path, zip_filepath, temp_path, store_path, 'ppt')

    elif str(path)[-5:] == '.xlsx':
        zip_filepath, store_path = makestorepath(sys.argv[1])
        office2pic(path, zip_filepath, temp_path, store_path, 'xl')

    # 暂时不知道旧格式用什么方式储存文件，总之不是简单压缩包。这里先做不支持处理
    elif str(path)[-4:] == '.doc' or str(path)[-4:] == '.ppt' or str(path)[-4:] == '.xls':
        alert(error('ErrorOldOfficeFormat') + str(str(path)[-4:]) + 'x' + error('ErrorOldOfficeFormatSuffix'),
              error('ErrorOnLoadTitle'))
        quit()

    else:
        alert(error('ErrorOnLoad'), error('ErrorOnLoadTitle'))
        quit()
