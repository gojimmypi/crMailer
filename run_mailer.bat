@echo off
:: Note crMailer needs [List Folder / Read Data] permissions on C:\Windows\temp\
:: but we'll try to set TMP to someplace else known to have permissions (not fully validated this works)
SET TEMP=d:\webdata\COVID19\temp
SET TMP=d:\webdata\COVID19\temp

:: the default is to start in C:\Windows\system32; we need to ensure we start in d:\tools\COVID19
d:
cd \tools\COVID19
crMailer.exe 
