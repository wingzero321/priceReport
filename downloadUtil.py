#!/usr/bin/env python
# -*- coding: utf-8 -*-

import paramiko
import xlsxFormatSetting as settings


server_ip = settings.server_ip
server_user = settings.server_user
server_passwd = settings.server_passwd
server_port = settings.server_port

def ssh_connect():
  ssh = paramiko.SSHClient()
  ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
  ssh.connect(server_ip, server_port,server_user, server_passwd)
  return ssh

def ssh_disconnect(client):
  client.close()

def exec_cmd(ssh,command):
  '''
  windows客户端远程执行linux服务器上命令
  '''
  stdin, stdout, stderr = ssh.exec_command(command)
  err = stderr.readline()
  out = stdout.readline()

  print command

  # if "" != err:
  #   print "command: " + command + " exec failed!\nERROR :" + err
  #   #return true, err
  # else:
  #   print "command: " + command + " exec success."
  #   print out

def win_to_linux(localpath, remotepath):
  '''
  windows向linux服务器上传文件.
  localpath  为本地文件的绝对路径。如：D:\test.py
  remotepath 为服务器端存放上传文件的绝对路径,而不是一个目录。如：/tmp/my_file.txt
  '''
  client = paramiko.Transport((server_ip, server_port))
  client.connect(username = server_user, password = server_passwd)
  sftp = paramiko.SFTPClient.from_transport(client)

  sftp.put(localpath,remotepath)
  client.close()

def linux_to_win(localpath, remotepath):
  '''
  从linux服务器下载文件到本地
  localpath  为本地文件的绝对路径。如：D:\test.py
  remotepath 为服务器端存放上传文件的绝对路径,而不是一个目录。如：/tmp/my_file.txt
  '''
  client = paramiko.Transport((server_ip, server_port))
  client.connect(username = server_user, password = server_passwd)
  sftp = paramiko.SFTPClient.from_transport(client)

  sftp.get(remotepath, localpath)
  client.close()

def download(filename):
    #windows环境文件路径
    winPath = settings.filePath + '\\data\\' + filename
    #linux环境文件路径
    linPath = settings.linPath + '/' + filename
    linux_to_win(winPath, linPath)