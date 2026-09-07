@echo off
chcp 65001 >nul
echo ================================================
echo   第一步: 启动 NapCat 并扫码登录 QQ
echo ================================================
echo.
echo 即将启动 NapCat,请用手机QQ扫描二维码图片完成登录
echo (二维码图片保存在: %USERPROFILE%\NapCatShell\cache\qrcode.png)
echo.
echo 登录成功后,运行「第2步_一键导出.bat」
echo.
pause
start "" /D "%USERPROFILE%\NapCatShell\bootmain" "%USERPROFILE%\NapCatShell\bootmain\QQ.exe" --enable-logging
