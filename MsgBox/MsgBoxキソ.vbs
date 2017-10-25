Option Explicit

'MsgBox関数のこと
'	2002/01/06

'文字を書いた方がわかりやすい。
'それと、上手に足し算するとかっこいい。

' 第二引数
' vbOKOnly				   0	[OK] ボタンのみを表示します。
' vbOKCancel			   1	[OK] ボタンと [キャンセル] ボタンを表示します。
' vbAbortRetryIgnore	   2	[中止]、[再試行]、および [無視] の 3 つのボタンを表示します。
' vbYesNoCancel			   3	[はい]、[いいえ]、および [キャンセル] の 3 つのボタンを表示します。
' vbYesNo				   4	[はい] ボタンと [いいえ] ボタンを表示します。
' vbRetryCancel			   5	[再試行] ボタンと [キャンセル] ボタンを表示します。
' vbCritical			  16	警告メッセージ アイコンを表示します。
' vbQuestion			  32	問い合わせメッセージ アイコンを表示します。
' vbExclamation			  48	注意メッセージ アイコンを表示します。
' vbInformation			  64	情報メッセージ アイコンを表示します。
' vbDefaultButton1		   0	第 1 ボタンを標準ボタンにします。
' vbDefaultButton2		 256	第 2 ボタンを標準ボタンにします。
' vbDefaultButton3		 512	第 3 ボタンを標準ボタンにします。
' vbDefaultButton4		 768	第 4 ボタンを標準ボタンにします。
' vbApplicationModal	   0	アプリケーション モーダルに設定します。メッセージ ボックスに応答するまで、現在選択中のアプリケーションの実行を継続できません。
' vbSystemModal			4096	システム モーダルに設定します。メッセージ ボックスに応答するまで、すべてのアプリケーションが中断されます。





'* 変数定義 *
Dim nRtn		' Number型の変数

'* 関数呼び出し *
nRtn = MsgBox("prompt!!", vbYesNoCancel + vbCritical, "title!!")

'* 戻り値表示 *
MsgBox(nRtn)

' 戻り値
' 定数     値 選択されたボタン 
' vbOK     1  [OK] 
' vbCancel 2  [キャンセル] 
' vbAbort  3  [中止] 
' vbRetry  4  [再試行] 
' vbIgnore 5  [無視] 
' vbYes    6  [はい] 
' vbNo     7  [いいえ] 

