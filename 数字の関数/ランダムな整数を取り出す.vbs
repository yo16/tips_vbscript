Dim MyValue, Response
Randomize   ' �����W�F�l���[�^�����������܂��B
Do Until Response = vbNo
   MyValue = Int((6 * Rnd) + 1)   ' 1 �` 6 �̃����_���Ȓl�𐶐����܂��B
   MsgBox MyValue
   Response = MsgBox ("�J��Ԃ��܂��� ? ", vbYesNo)
Loop



