


set REG_PATH=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set EXP_FILE=Export.txt


reg export %REG_PATH% %EXP_FILE%

echo ‚Å‚«‚½ƒtƒ@ƒCƒ‹‚ÍUTF16
pause
