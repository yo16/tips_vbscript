Option Explicit


Dim array1
array1 = Array("a","b","c")
' ���̌��ɗv�f�������P�t�����������Ȃ��B�B

'*****�����B*****
' �z��𒷂��Ē�`(�l��ێ��������Ƃ���Preserve���w�肷��)
ReDim Preserve array1(UBound(array1)+1)
' �z��̍Ō�ɗv�f��ǉ�
array1(UBound(array1)) = "x"

msgbox array1(0)
msgbox array1(1)
msgbox array1(2)
msgbox array1(3)

