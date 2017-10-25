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
'--�v���p�e�B�͂��ׂ�Boolean�^
	Private insObjName		'--�t�@�C����
	Private insReadOnly		'--�ǂݎ���p(I/O)
	Private insHidden		'--�B���t�@�C��(I/O)
	Private insSystem		'--�V�X�e���t�@�C��(I/O)
	Private insDirectory		'--�t�H���_(O)
	Private insArchive		'--�A�[�J�C�u(I/O)
	Private insAlias		'--�V���[�g�J�b�g(O)
	Private insCompressed		'--���k(O)

	Private objType			'--�t�@�C���̏ꍇ"File"�B�t�H���_�̏ꍇ"Folder"�B�i�����ϐ�)
	Private propertyDictionary	'--�v���p�e�B�Ɛ��l�̑Ή�


'*************
'Property��Get
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
'Property��Let
'*************
	Public Property Let FileName(newProperty)
		Dim objFS
		Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
		If Not objFS.FileExists(newProperty) Then
			If Not objFS.FolderExists(newProperty) Then
				MsgBox newProperty&"�Ƃ����t�@�C�����t�H���_�����݂��܂���I"
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
'�O���֐�
'********
	'�� �v���p�e�B�ꗗ��\������
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
'�����֐�
'********
	'�� FileName��Let����Ă��邩�`�F�b�N����
	Private Sub checkSetFile
		If (insObjName = "") Then
			MsgBox "Property[FileName] ���w�肳��Ă��܂���I"
			WScript.Quit
		End If
	End Sub

	'�� ���������������`�F�b�N����
	'		pValue	:�^���`�F�b�N����l
	Private Sub checkPropertyValue(pValue)
		If (TypeName(pValue) <> "Boolean") Then
			MsgBox "True or False (Boolean�^)�̂ݐݒ�\�ł��I"
			WScript.Quit
		End If
	End Sub

	'�� ���ׂẴv���p�e�B���N���A����
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

	'�� �v���p�e�B�r�b�g�̊֘A�z������
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

	'�� �v���p�e�B���Z�b�g����
	'		propertyName[String]	:�v���p�e�B�̖��O(Dictionary�Q��)
	'		propertyValue[Boolean]	:�Z�b�g����l
	Private Sub setPropertyBit(propertyName,propertyValue)
		Dim objFS,objF
		Set objFS = CreateObject("Scripting.FileSystemObject")
		If (objType = "File") Then
			Set objF = objFS.GetFile(insObjName)
		ElseIf (objType = "Folder") Then
			Set objF = objFS.GetFolder(insObjName)
		Else
			MsgBox "objType�����܂��ĂȂ��悤�ł��I"
			WScript.Quit
		End If

		Dim bitValue
		bitValue = propertyDictionary.Item(propertyName)
		If (objF.Attributes and bitValue) Then	'--True�̏ꍇ
			If Not propertyValue Then	'--False
				objF.Attributes = objF.Attributes - bitValue
			End If
		Else									'--False�̏ꍇ
			If propertyValue Then	'--True
				objF.Attributes = objF.Attributes + bitValue
			End If
		End If
	End Sub

'**************
'�n�܂�ƏI���
'**************
	Private Sub Class_Initialize
		allPropertyClear
		makeDictionary
	End Sub
	Private Sub Class_Terminate
	End Sub

End Class
