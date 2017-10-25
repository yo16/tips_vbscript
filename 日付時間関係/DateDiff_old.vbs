Option Explicit
'-----------------------------------------------------
' DataDiffのテスト
'-----------------------------------------------------

' 今日
Dim v_Today
v_Today = Date

' 指定日
Dim v_Aruhi
v_Aruhi = CDate("2006/5/30")


Dim v_diffDays
v_diffDays = DateDiff( "y", v_Today, v_Aruhi )
' １つ目の引数
' 設定値	内容 
' yyyy		年 
' q			四半期 
' m			月 
' y			年間通算日 
' d			日 
' w			週日 
' ww		週 
' h			時 
' n			分 
' s			秒 





msgbox( v_diffDays )

