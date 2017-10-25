
set DNS1=10.32.89.0
set DNS2=10.32.89.1

set DNS1=8.8.8.8
set DNS2=8.8.4.4

netsh interface ip set dns "ワイヤレス ネットワーク接続" static %DNS1% primary
netsh interface ip add dns "ワイヤレス ネットワーク接続" %DNS2%
pause
