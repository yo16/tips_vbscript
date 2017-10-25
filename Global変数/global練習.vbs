
Dim X         ' グローバル スコープで X を宣言します。
X = "Global"      ' グローバルな X に値を代入します。

Sub Proc1   ' プロシージャを宣言します。
Dim X      ' ローカル スコープで X を宣言します。
X = "Local"   ' ローカルな X に値を代入します。
         ' 呼び出されると X を出力するプロシージャを、
         ' Execute ステートメントで作成します。
         ' グローバル スコープに含まれるすべてを Proc2 が
         ' 継承するため、グローバルな X が出力されます。
  ExecuteGlobal "Sub Proc2: Print X: End Sub"
Print Eval("X")   ' ローカルな X を出力します。


Proc2      ' グローバル スコープで Proc2 を呼び出すと、
         ' "Global" が印刷されます。
End Sub

Proc2         ' Proc1 の外部で Proc2 を使用できないため、
         ' この行でエラーが発生します。
Proc1         ' Proc1 を呼び出します。
  Execute "Sub Proc2: Print X: End Sub"
Proc2         ' Proc2 をグローバルに使用できるように
         ' なったので、この呼び出しは成功します。


