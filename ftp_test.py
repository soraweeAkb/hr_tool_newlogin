import ftplib
import os
from pathlib import Path
import time


def ftpDownload(curTime):
    host = r'210.1.31.3'
    #port = 21
    user = 'lu'
    password = 'B*c913ke9'
    LocalDir = 'D:\HR_TOOL_UPDATE20221021\source\main.exe'
    RemoteDir = '/HR_tool_V2.0.exe'
    ftp = ftplib.FTP()
    ftp.connect(host)
    ftp.login(user, password)
    DownloadFile(LocalDir, RemoteDir, ftp)
    ftp.close()
    print("文件下载完成")


# 下载单个文件
def DownloadFile(LocalFile, RemoteFile, ftp):
    file_handler = open(LocalFile, 'wb')
    print(file_handler)
    print('----------',RemoteFile)
    ftp.retrbinary('RETR '+RemoteFile, file_handler.write)
    file_handler.close()
    return True


if __name__ == '__main__':
    curTime = time.strftime("%Y-%m-%d_%H-%M")

    file = ftpDownload(curTime)
