
set DNS1=10.32.89.0
set DNS2=10.32.89.1

set DNS1=8.8.8.8
set DNS2=8.8.4.4

netsh interface ip set dns "���C�����X �l�b�g���[�N�ڑ�" static %DNS1% primary
netsh interface ip add dns "���C�����X �l�b�g���[�N�ڑ�" %DNS2%
pause
