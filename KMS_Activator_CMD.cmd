@echo off
title KMS 激活工具（保持联网永久激活）
color 0c
#=====================================
setlocal EnableDelayedExpansion& cd /d "%~dp0"
#=====================================
mode con cols=35 lines=10
cls
echo   KMS 激活工具（保持联网永久激活）
echo.
echo           选择激活服务器
echo.
echo 1、零散坑の服务器（默认）
echo 2、kms.pubの服务器
echo.
echo.
set /p user_input_host=请选择要激活的项目：
if not defined user_input_host goto:default_host
if %user_input_host% equ 1 set kms_host=kms.03k.org
if %user_input_host% equ 2 set kms_host=win.kms.pub
goto mainmenu
:default_host
set kms_host=kms.03k.org
goto mainmenu
#======================================
:mainmenu
set user_input_main=
mode con cols=35 lines=12
cls
echo   KMS 激活工具（保持联网永久激活）
echo.
echo                主菜单
echo.
echo 1、Windows 激活
echo 2、Office 激活
echo Q、退出程序
echo.
echo                      版本12_230619
echo.
set /p user_input_main=请选择要激活的项目：
if not defined user_input_main goto:mainmenu
if %user_input_main% equ 1 goto menu
if %user_input_main% equ 2 goto office
if %user_input_main% equ q exit
if %user_input_main% equ Q exit
#=====================================
:office
set user_input_o=
mode con cols=35 lines=8
cls
echo.
echo   KMS 激活工具（保持联网永久激活）
echo.
echo   V12 版本开始不再提供Office激活
echo   推荐使用 Office Tool Plus 激活
echo.
set /p user_input_o=是否打开OTP网站(Y/N,默认Y)：
if not defined user_input_o goto:otp
if %user_input_o% equ N goto mainmenu
if %user_input_o% equ n goto mainmenu
if %user_input_o% equ Y goto otp
if %user_input_o% equ y goto otp
goto mainmenu
:otp
start https://otp.landian.vip/zh-cn/
goto mainmenu
#=====================================
:menu
set user_input_os=
mode con cols=35 lines=19
cls
echo   KMS 激活工具（保持联网永久激活）
echo.
echo            Windows 激活
echo.
echo 1、专业版
echo 2、专业工作站版
echo 3、专业教育版
echo 4、教育版
echo 5、企业版
echo C、检测系统版本（Beta）
echo U、更新KMS激活服务器
echo Q、返回主菜单
echo.
echo *注：选项U为仅更新KMS服务器地址
echo ..请确认上次激活使用了KMS
echo ..否则本项不能使用
echo.
set /p user_input_os=请选择系统版本：
if not defined user_input_os goto:menu
if %user_input_os% equ 1 goto pro
if %user_input_os% equ 2 goto prostation
if %user_input_os% equ 3 goto prolearn
if %user_input_os% equ 4 goto learn
if %user_input_os% equ 5 goto ep
if %user_input_os% equ c goto check_os
if %user_input_os% equ C goto check_os
if %user_input_os% equ u goto finish
if %user_input_os% equ U goto finish
if %user_input_os% equ q goto mainmenu
if %user_input_os% equ Q goto mainmenu
:check_os
mode con cols=35 lines=8
cls
echo.
echo      当前系统版本（Beta）
echo.
wmic os get caption
pause
goto menu
#=====================================
:finish
slmgr /skms %kms_host%
slmgr /ato
goto mainmenu

#10专业版
:pro
slmgr /upk
slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
goto finish

#10专业工作站版
:prostation
slmgr /upk
slmgr /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
goto finish

#10专业教育版
:prolearn
slmgr /upk
slmgr /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
goto finish

#10教育版
:learn
slmgr /upk
slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
goto finish

#10企业版
:ep
slmgr /upk
slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
goto finish
