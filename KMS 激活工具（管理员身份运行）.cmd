@echo off
title KMS 激活工具（保持联网永久激活）
color 0c
#=====================================
setlocal EnableDelayedExpansion& cd /d "%~dp0"
#=====================================
:mainmenu
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
echo                      版本10_210524
echo.
set /p user_input_main=请选择要激活的项目：
if not defined user_input_main goto:erra
if %user_input_main% equ 1 goto menu
if %user_input_main% equ 2 goto office
if %user_input_main% equ q exit
if %user_input_main% equ Q exit

:erra
goto mainmenu
#=====================================

:office
mode con cols=60 lines=22

:runas
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

cls
echo                KMS 激活工具（保持联网永久激活）
echo.
echo                        Office 激活
echo.
echo 1、Office Pro Plus 2016
echo 2、Office Visio Pro 2016
echo 3、Office Project Pro 2016
echo 4、Office Pro Plus 2019
echo 5、Office Visio Pro 2019
echo 6、Office Project Pro 2019
echo 7、Office Pro Plus 2021(beta)
echo 8、Office Visio Pro 2021(beta)
echo 9、Office Project Pro 2021(beta)
echo U、更新KMS激活服务器
echo Q、返回主菜单
echo.
echo *注：选项U为仅更新KMS服务器地址
echo ..请确认上次激活使用了KMS
echo ..否则本项不能使用
echo.
set /p user_input_office=请选择Office版本：
if not defined user_input_office goto:errb
if %user_input_office% equ 1 goto opp16
if %user_input_office% equ 2 goto ovp16
if %user_input_office% equ 3 goto opj16
if %user_input_office% equ 4 goto opp19
if %user_input_office% equ 5 goto ovp19
if %user_input_office% equ 6 goto opj19
if %user_input_office% equ 7 goto opp21
if %user_input_office% equ 8 goto ovp21
if %user_input_office% equ 9 goto opj21
if %user_input_office% equ u goto e
if %user_input_office% equ U goto e
if %user_input_office% equ q goto mainmenu
if %user_input_office% equ Q goto mainmenu
:errb
goto mainmenu
#=====================================
:opp16
 
cls
 
echo 正在重置Office2016零售激活...
cscript ospp.vbs /rearm
 
echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
 
echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
 
echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
 
goto e
#=====================================
:ovp16
cls
 
echo 正在重置Visio2016零售激活...
cscript ospp.vbs /rearm
 
echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
 
echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
 
echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
 
goto e
#=====================================
:opj16

cls
 
echo 正在重置Project2016零售激活...
cscript ospp.vbs /rearm
 
echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
 
echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
 
echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT
 
goto e
#=====================================
:opp19

cls

echo 正在重置Office2019零售激活...
cscript ospp.vbs /rearm

echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP

goto e
#=====================================
:ovp19
cls

echo 正在重置Visio2019零售激活...
cscript ospp.vbs /rearm

echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB

goto :e
#=====================================
:opj19

cls

echo 正在重置Project2019零售激活...
cscript ospp.vbs /rearm

echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B

goto :e
#=====================================
:opp21

cls

echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:HFPBN-RYGG8-HQWCW-26CH6-PDPVF

goto e
#=====================================
:ovp21
cls

echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:2XYX7-NXXBK-9CK7W-K2TKW-JFJ7G

goto :e
#=====================================
:opj21

cls

echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:WDNBY-PCYFY-9WP6G-BXVXM-92HDV

goto :e
#=====================================
:e
cscript ospp.vbs /sethst:kms.03k.org
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo 激活流程结束
pause
goto mainmenu
#=====================================
:menu
mode con cols=35 lines=19
cls
echo   KMS 激活工具（保持联网永久激活）
echo.
echo          Windows 10 激活
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
if not defined user_input_os goto:errc
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
:errc
goto menu
#=====================================
:finish
slmgr /skms kms.03k.org
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