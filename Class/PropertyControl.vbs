Option Explicit

Dim objPC
Set objPC = New PropertyControl

objPC.FileName = "a"
'MSGBOX objPC.ReadOnly,,"ReadOnly"
'objPC.ReadOnly = True
'MSGBOX objPC.ReadOnly,,"ReadOnly"
'MSGBOX objPC.Hidden,,"Hidden"
'objPC.Hidden = True
'MSGBOX objPC.Hidden,,"Hidden"

MsgBox objPC.AllPropertys,,"All Propertys"



Class PropertyControl
'--プロパティはすべてBoolean型
	Private insObjName		'--ファイル名
	Private insReadOnly		'--読み取り専用(I/O)
	Private insHidden		'--隠しファイル(I/O)
	Private insSystem		'--システムファイル(I/O)
	Private insDirectory		'--フォルダ(O)
	Private insArchive		'--アーカイブ(I/O)
	Private insAlias		'--ショートカット(O)
	Private insCompressed		'--圧縮(O)

	Private objType			'--ファイルの場合"File"。フォルダの場合"Folder"。（内部変数)
	Private propertyDictionary	'--プロパティと数値の対応


'*************
'PropertyをGet
'*************
	Public Property Get ReadOnly
		checkSetFile
		ReadOnly = insReadOnly
	End Property
	Public Property Get Hidden
		checkSetFile
		Hidden = insHidden
	End Property
	Public Property Get System
		checkSetFile
		System = insSystem
	End Property
	Public Property Get Directory
		checkSetFile
		Directory = insDirectory
	End Property
	Public Property Get Archive
		checkSetFile
		Archive = insArchive
	End Property
	Public Property Get Alias
		checkSetFile
		Alias = insAlias
	End Property
	Public Property Get Compressed
		checkSetFile
		Compressed = insCompressed
	End Property


'*************
'PropertyをLet
'*************
	Public Property Let FileName(newProperty)
		Dim objFS
		Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
		If Not objFS.FileExists(newProperty) Then
			If Not objFS.FolderExists(newProperty) Then
				MsgBox newProperty&"というファイルもフォルダも存在しません！"
				WScript.Quit
			Else
				objType = "Folder"
			End If
		Else
			objType = "File"
		End If
		allPropertyClear
		insObjName = newProperty

		Dim objFile
		If (objType = "File") Then
			Set objFile = objFS.GetFile(newProperty)
		ElseIf (objType = "Folder") Then
			Set objFile = objFS.GetFolder(newProperty)
		End If

		If objFile.attributes and propertyDictionary.Item("ReadOnly") Then	'--ReadOnly
			insReadOnly = True
		Else
			insReadOnly = False
		End If
		If objFile.attributes and propertyDictionary.Item("Hidden") Then	'--Hidden
			insHidden = True
		Else
			insHidden = False
		End If
		If objFile.attributes and propertyDictionary.Item("System") Then	'--Sysmtem
			insSystem = True
		Else
			insSystem = False
		End If
		If objFile.attributes and propertyDictionary.Item("Directory") Then	'--Directory
			insDirectory = True
		Else
			insDirectory = False
		End If
		If objFile.attributes and propertyDictionary.Item("Archive") Then	'--Archive
			insArchive = True
		Else
			insArchive = False
		End If
		If objFile.attributes and propertyDictionary.Item("Alias") Then	'--Alias
			insAlias = True
		Else
			insAlias = False
		End If
		If objFile.attributes and propertyDictionary.Item("Compressed") Then	'--Compressed
			insCompressed = True
		Else
			insCompressed = False
		End If
	End Property

	Public Property Let ReadOnly(newProperty)
		checkSetFile
		checkPropertyValue(newProperty)
		setPropertyBit "ReadOnly",newProperty
		insReadOnly = newProperty
	End Property

	Public Property Let Hidden(newProperty)
		checkSetFile
		checkPropertyValue(newProperty)
		setPropertyBit "Hidden",newProperty
		insHidden = newProperty
	End Property

	Public Property Let System(newProperty)
		checkSetFile
		checkPropertyValue(newProperty)
		setPropertyBit "System",newProperty
		insSystem = newProperty
	End Property

	Public Property Let Archive(newProperty)
		checkSetFile
		checkPropertyValue(newProperty)
		setPropertyBit "Archive",newProperty
		insArchive = newProperty
	End Property


'********
'外部関数
'********
	'＊ プロパティ一覧を表示する
	Public Function AllPropertys
		checkSetFile
		Dim msgStr
		msgStr =	"ReadOnly : "&insReadOnly&vbCrLf&_
					"Hidden : "&insHidden&vbCrLf&_
					"System : "&insSystem&vbCrLf&_
					"Directory : "&insDirectory&vbCrLf&_
					"Archive : "&insArchive&vbCrLf&_
					"Alias : "&insAlias&vbCrLf&_
					"Compressed : "&insCompressed
		AllPropertys = msgStr
	End Function


'********
'内部関数
'********
	'＊ FileNameがLetされているかチェックする
	Private Sub checkSetFile
		If (insObjName = "") Then
			MsgBox "Property[FileName] が指定されていません！"
			WScript.Quit
		End If
	End Sub

	'＊ 引数が正しいかチェックする
	'		pValue	:型をチェックする値
	Private Sub checkPropertyValue(pValue)
		If (TypeName(pValue) <> "Boolean") Then
			MsgBox "True or False (Boolean型)のみ設定可能です！"
			WScript.Quit
		End If
	End Sub

	'＊ すべてのプロパティをクリアする
	Private Sub allPropertyClear
		insObjName = ""
		insReadOnly = ""
		insHidden = ""
		insSystem = ""
		insDirectory = ""
		insArchive = ""
		insAlias = ""
		insCompressed = ""
	End Sub

	'＊ プロパティビットの関連配列を作る
	Private Sub makeDictionary
		Set propertyDictionary = CreateObject("Scripting.Dictionary")
		propertyDictionary.CompareMode = vbTextCompare
		propertyDictionary.Add "Normal","0"
		propertyDictionary.Add "ReadOnly","1"
		propertyDictionary.Add "Hidden","2"
		propertyDictionary.Add "System","4"
		propertyDictionary.Add "Directory","16"
		propertyDictionary.Add "Archive","32"
		propertyDictionary.Add "Alias","1024"
		propertyDictionary.Add "Compressed","2048"
	End Sub

	'＊ プロパティをセットする
	'		propertyName[String]	:プロパティの名前(Dictionary参照)
	'		propertyValue[Boolean]	:セットする値
	Private Sub setPropertyBit(propertyName,propertyValue)
		Dim objFS,objF
		Set objFS = CreateObject("Scripting.FileSystemObject")
		If (objType = "File") Then
			Set objF = objFS.GetFile(insObjName)
		ElseIf (objType = "Folder") Then
			Set objF = objFS.GetFolder(insObjName)
		Else
			MsgBox "objTypeが決まってないようです！"
			WScript.Quit
		End If

		Dim bitValue
		bitValue = propertyDictionary.Item(propertyName)
		If (objF.Attributes and bitValue) Then	'--Trueの場合
			If Not propertyValue Then	'--False
				objF.Attributes = objF.Attributes - bitValue
			End If
		Else									'--Falseの場合
			If propertyValue Then	'--True
				objF.Attributes = objF.Attributes + bitValue
			End If
		End If
	End Sub

'**************
'始まりと終わり
'**************
	Private Sub Class_Initialize
		allPropertyClear
		makeDictionary
	End Sub
	Private Sub Class_Terminate
	End Sub

End Class
