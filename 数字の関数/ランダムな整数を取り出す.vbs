Dim MyValue, Response
Randomize   ' 乱数ジェネレータを初期化します。
Do Until Response = vbNo
   MyValue = Int((6 * Rnd) + 1)   ' 1 ～ 6 のランダムな値を生成します。
   MsgBox MyValue
   Response = MsgBox ("繰り返しますか ? ", vbYesNo)
Loop



