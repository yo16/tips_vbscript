' �[������ł����܂������Ă邩�H

Option Explicit


Dim array1
array1 = Array()

' msgbox UBound(array1)

'*****�����B*****
' �z��𒷂��Ē�`(�l��ێ��������Ƃ���Preserve���w�肷��)
ReDim Preserve array1(UBound(array1)+1)
' �z��̍Ō�ɗv�f��ǉ�
array1(UBound(array1)) = "x"



' �z��𒷂��Ē�`(�l��ێ��������Ƃ���Preserve���w�肷��)
ReDim Preserve array1(UBound(array1)+1)
' �z��̍Ō�ɗv�f��ǉ�
array1(UBound(array1)) = "y"



' �z��𒷂��Ē�`(�l��ێ��������Ƃ���Preserve���w�肷��)
ReDim Preserve array1(UBound(array1)+1)
' �z��̍Ō�ɗv�f��ǉ�
array1(UBound(array1)) = "z"

msgbox array1(0)
msgbox array1(1)
msgbox array1(2)
