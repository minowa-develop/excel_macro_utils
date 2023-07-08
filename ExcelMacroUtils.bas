'-------------------------------
'-- ExcelMacroUtils ver 0.2.1 --
'-------------------------------

'== Functions ==
'regex(�Ώە���t, �u��������(���K�\���g�p�\)p, �u������r)
'isMatchRegex(�Ώە���t, �u��������(���K�\���g�p�\)p)
'IsBlankRide(�`�F�b�N���镶��)
'fileNameIncrement(�t�@�C����(�g���q����, �h�b�g��������), �`�F�b�N����p�X)
'getSearchRowNo(�����l, �V�[�g�ԍ�, �J�E���g�J�n�s, �J�E���g��)
'SWITCH(�����l, �팟���l1, �擾�l1, �팟���l2, �擾�l2, �c, �l���Ȃ��ꍇ�̕\��)
'ConvertRange(�n�__��ԍ�, �n�__�s�ԍ�, �I�__��ԍ�, �I�__�s�ԍ�)
'ConvertColCode(��ԍ�)
'ConvertColNo(��R�[�h)
'VarticalValueGet(�V�[�g�ԍ�,�擾�J�n�s�ԍ�,��ԍ�)
'VarticalValueGetAdd(�V�[�g�ԍ�,�擾�J�n�s�ԍ�,��ԍ�,�ǉ���collection)
'maxRowCount(�V�[�g�ԍ�,�J�E���g�J�n�s�ԍ�,��ԍ�)
'ExactConmaValues(�����l, �팟���l(�R���}��؂�))
'trimRide(������)
'getnbsp()
'sdel(������)
'getFiles(�p�X,�g���q)
'ArrayToCollection(�ϊ��O(Collection))
'CollectionToArray(�ϊ��O(array))
'TEXTJOIN(��؂蕶��,�󔒂𖳎����邩�ǂ���,join�Ώۂ̕���...)

'== sub ==
'FileOut(�t�@�C����, ���e(��s����), �o�͂���p�X)
'autofill(�I�[�g�t�B���̑ΏۂƂȂ�Z��, �I�[�g�t�B���ő��₷�s��)
'ClearSheet(�V�[�g�ԍ�)
'deactivateDispUpdate
'activateDispUpdate


Option Explicit

' @title regex
' @param �Ώە���t, �u��������(���K�\���g�p�\)p, �u������r
' @sinse 2018.04.17
' @memo ���K�\�����g����u���֐��ł��B
' @help https://www.sejuku.net/blog/33541
Function regex(target As String, ptn As String, rep As String) as String 
  'RegExp�I�u�W�F�N�g�̍쐬
  Dim REG As Object: Set REG = CreateObject("VBScript.RegExp")

  '���K�\���̎w��
  With REG
    .PATTERN = ptn         '�p�^�[�����w��
    .IgnoreCase = False     '�啶���Ə���������ʂ��邩(False)�A���Ȃ���(True)
    .Global = True          '������S�̂��������邩(True)�A���Ȃ���(False)
  End With

  regex = REG.replace(target, rep) '�w�肵�����K�\�����2�����̋�؂蕶���ɒu������
End Function

' @title isMatchRegex
' @param �Ώە���t, �u��������(���K�\���g�p�\)p, �u������r
' @sinse 2018.04.17
' @memo ���K�\���Ƀ}�b�`������true��Ԃ�
Function isMatchRegex(target As String, ptn As String) As Boolean
  'RegExp�I�u�W�F�N�g�̍쐬
  Dim REG As Object: Set REG = CreateObject("VBScript.RegExp")

  '���K�\���̎w��
  With REG
    .PATTERN = ptn         '�p�^�[�����w��
    .IgnoreCase = False     '�啶���Ə���������ʂ��邩(False)�A���Ȃ���(True)
    .Global = True          '������S�̂��������邩(True)�A���Ȃ���(False)
  End With

  isMatchRegex = REG.Test(target) '�w�肵�����K�\�����2�����̋�؂蕶���ɒu������
End Function


' @title IsBlankRide
' @param �`�F�b�N���镶��
' @sinse 2018.04.17
' @update 2019.07.08
' @memo �󕶎��Ƌ󔒕����̏ꍇtrue���������܂��B
'       ������ISBLANK�͋󕶎��������Ȃ��Ȃ񂿂���ĂԂ�񂭂Ȃ̂ł��イ��
'       nbsp�Ή�
Function IsBlankRide(str As String) as String 
  '�萔
  Dim nbsp As String: nbsp = getnbsp()
  Dim REG As String: REG = "[\s" + nbsp + "]"

  Dim return_flg As Boolean: return_flg = False

  '�󔒍폜
  Dim strRep As String
  strRep = regex(str, REG, "")

  '����
  If strRep = "" Then return_flg = True

  IsBlankRide = return_flg
End Function


' @title FileNameIncrement
' @param �t�@�C����(�g���q����, �h�b�g��������), �`�F�b�N����p�X
' @sinse 2018.05.10
' @memo �����̃t�@�C�����������ꍇ(i)�`���ɂ���
Function FileNameIncrement(fileName As String, path As String) as String
  '�����`�F�b�N
  If Dir(path & "\" + fileName) <> "" Then
    '�C���N�������g�t�^
    ' �Ō��.���q�b�g������ɂ� p:\.([^\.]+)$ �ŃL���v�`���������g���q�ƂȂ�
    ' fileName = regex(fileName, "\.", "(1).")
    fileName = regex(fileName, "\.([^.]+$)", "(1).$1")
    Dim i As Integer: i = 1
    Do While Dir(path & "\" + fileName) <> ""
      '�C���N�������g
      fileName = regex(fileName, "\([0-9]+\)", "(" + CStr(i) + ")")
      i = i + 1
    Loop
  End If

  FileNameIncrement = fileName
End Function


' @title getSearchRowNo
' @param �����l,�V�[�g�ԍ�,�J�E���g�J�n�s,�J�E���g��
' @sinse 2019.01.30
' @memo  �����Ƀq�b�g�����s��(�J�E���g�J�n�s����J�E���g)��Ԃ��B
'        �q�b�g���Ȃ��ꍇ��-1��Ԃ�
Function getSearchRowNo(target As String, sheetNo As Integer, startRow As Integer, col As Integer) as long
  Dim sheet As Worksheet: Set sheet = ThisWorkbook.Worksheets(sheetNo)
  Dim i As Integer: i = startRow
  Dim hitRow As Integer: hitRow = -1

  Do While sheet.Cells(i, col).value <> "" And hitRow = -1
    If target = sheet.Cells(i, col).value Then
      hitRow = i - startRow + 1
    End If
    i = i + 1
  Loop

  getSearchRowNo = hitRow
End Function


' @title SWITCH
' @sinse 2019.03.14
' @memo excel2010��switch�֐����g���Ȃ��̂ő�p�B
'       �Q�l�R�s�y��(https://www.excelspeedup.com/switch/)
' @memo (?<key>[^\t]+)\t(?<value>[^\r\n]+)\r\n �� "\k<key>","\k<value>",
' @memo SWITCH(search,key1,value1,key2,value2,�c,else_value)
Function SWITCH(ParamArray par())
  Dim i As Integer

  For i = LBound(par) + 1 To UBound(par) - 1 Step 2
    If par(LBound(par)) = par(i) Then
      SWITCH = par(i + 1)
      Exit Function
    End If
  Next

  If i = UBound(par) Then
    SWITCH = par(i)
  Else
    SWITCH = CVErr(xlErrNA)
  End If
End Function


' @title ConvertRange
' @param �n�__��ԍ�, �n�__�s�ԍ�, �I�__��ԍ�, �I�__�s�ԍ�
' @sinse 2019.05.16
' @memo Range�`���ɕϊ����� ��)A1:B3
Function ConvertRange(colNo_from As Long, rowNo_from As Long, colNo_to As Long, rowNo_to As Long) as String 
  Dim rangeBuf As String: rangeBuf = ConvertColCode(colNo_from) & rowNo_from & ":" & ConvertColCode(colNo_to) & rowNo_to
  ConvertRange = rangeBuf
End Function


' @title ConvertColCode
' @param ��ԍ�
' @sinse 2019.05.16
' @memo ��ԍ������R�[�h�ɕϊ�
Function ConvertColCode(colNo As Long) as String 
  Dim buf As String
  buf = Cells(1, colNo).Address(True, False)
  buf = Left(buf, InStr(buf, "$") - 1)

  ConvertColCode = buf
End Function


' @title ConvertColNo
' @param ��R�[�h
' @sinse 2020.07.15
' @memo ��R�[�h�����ԍ��ɕϊ�
Function ConvertColNo(colCode As String) As Long
  ConvertColNo = Range(colCode & "1").Column
End Function


' @title ExactConma
' @param �����l, �팟���l(�R���}��؂�)
' @sinse 2019.05.31
' @memo �팟���l�R���}��؂肩�猟���l�������邩�T��
Function ExactConmaValues(search As String, str As String) as boolean
  Dim ret As Boolean: ret = False
  'conmaconvert
  Dim strList As Variant: strList = Split(str, ",")

  Dim i As Integer: i = 0
  For i = 0 To UBound(strList)
    ret = search = strList(i)
      'tmp��true�̏ꍇfor����o����
      If ret Then Exit For
  Next i

  ExactConmaValues = ret
End Function


' @title trimRide
' @param ������
' @sinse 2019.07.03
' @memo �擪������\s��nbsp���폜����
'       trim�֐���nbsp�ɑΉ����Ă��Ȃ��B
Function trimRide(str As String) as String 
  Dim nbsp As String: nbsp = getnbsp()
  Dim ptn_module As String: ptn_module = "[\s" & nbsp & "]+"
  Dim ptn As String: ptn = "^" & ptn_module & "|" & ptn_module & "$"
  
  trimRide = regex(str, ptn, "")
End Function


' @title getnbsp
' @sinse 2019.07.03
' @memo nbsp�擾(no break space 00a0)
Function getnbsp()
  getnbsp = ChrW(160)
End Function


' \s�n����菜��
Function sDel(target As String) As String
  Dim nbsp As String: nbsp = getnbsp()
  Dim ptn As String: ptn = "[\s" & nbsp & vbCrLf & "�@]+"
  
  sDel = regex(target, ptn, "")
End Function


' @title getFiles
' @param �p�X,�g���q
' @sinse 2020.05.20
' @memo �w�肵���p�X���Ɏw�肵���g���q�̃t�@�C�������X�g�Ŏ擾����
Function getFiles(path As String, filterKakuthoushi As String) as collection 
    Dim buf As String, cnt As Long
    buf = Dir(path & "*." & filterKakuthoushi)
    Dim fileList As Collection: Set fileList = New Collection
    Do While buf <> ""
        cnt = cnt + 1
        fileList.Add (buf)
        buf = Dir()
    Loop
    
    Set getFiles = fileList
End Function


' @title ArrayToCollection
' @param �z��
' @sinse 2020.05.20
' @memo �z���collection�ɕϊ�����
' @memo �R�s�y�� https://oki2a24.com/2016/03/07/array-to-flat-collection-in-excel-vba/
Function ArrayToCollection(ByVal arr As Variant) As Collection
  Set ArrayToCollection = New Collection
  ' �������z��łȂ��ꍇ�͗v�f�� 1 �̃R���N�V�����Ƃ��ĕԂ��B
  If Not IsArray(arr) Then
    Call ArrayToCollection.Add(arr)
    Exit Function
  End If

  Dim v As Variant
  For Each v In arr
    If IsArray(v) Then
      ' �z��̗v�f���z��̏ꍇ�͍ċA�I�ɏ���
      Dim c As Collection: Set c = ArrayToCollection(v)
      Dim v2 As Variant
      For Each v2 In c
          Call ArrayToCollection.Add(v2)
      Next v2
    Else
      Call ArrayToCollection.Add(v)
    End If
  Next v
End Function


' @title ArrayToCollection
' @param collection
' @sinse 2020.05.20
' @memo collection��z��ɕϊ�����
' @memo �R�s�y�� https://oki2a24.com/2015/12/05/collection-to-array-with-excel-vba/
Function CollectionToArray(ByVal colTarget As Collection) As Variant
  Dim vntResult As Variant
  ReDim vntResult(colTarget.count - 1)
 
  Dim i As Long
  i = LBound(vntResult)
  Dim v As Variant
  For Each v In colTarget
    vntResult(i) = v
    i = i + 1
  Next v
 
  CollectionToArray = vntResult
End Function





' @title textjoin
' @param collection
' @sinse 2020.05.22
' @memo excel2016�Ŏg���Ȃ�textjoin���g����悤�ɂ���֐�
Function TEXTJOIN(Delim, Ignore As Boolean, ParamArray par()) as String 
  Dim i As Integer
  Dim tR As Range
   
  TEXTJOIN = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        If tR.value <> "" Or Ignore = False Then
          TEXTJOIN = TEXTJOIN & Delim & tR.Value2
        End If
      Next
    Else
      If par(i) <> "" Or Ignore = False Then
        TEXTJOIN = TEXTJOIN & Delim & par(i)
      End If
    End If
  Next
  
  TEXTJOIN = Mid(TEXTJOIN, Len(Delim) + 1)
End Function

' @title VarticalValueGet
' @param �V�[�g�ԍ�,�擾�J�n�s�ԍ�,��ԍ�
' @sinse 2018.04.27
' @memo �w���1����擾����
'       ��s������Ǝ擾���I������B
Function VarticalValueGet(sheetNo As Integer, beginRowNo As Long, colNo As Long) as collection
  Dim wk As Worksheet: Set wk = ThisWorkbook.Worksheets(sheetNo)
  Dim arrayValue As Collection: Set arrayValue = New Collection

  '�V�[�g���e�擾
  Dim curRowNo As Long: curRowNo = beginRowNo
  Do While wk.Cells(curRowNo, colNo).value <> ""
    arrayValue.Add (wk.Cells(curRowNo, colNo).value)
    curRowNo = curRowNo + 1
  Loop

  Set VarticalValueGet = arrayValue
End Function

' @title VarticalValueGet��������collection�ɒǉ�����֐�
' @param �V�[�g�ԍ�,�擾�J�n�s�ԍ�,��ԍ�,�ǉ���collection
' @sinse 2018.04.27
' @memo �w���1����擾����
'       ��s������Ǝ擾���I������B
Function VarticalValueGetAdd(sheetNo As Integer, beginRowNo As Long, colNo As Long, list As Collection) as collection
  Dim wk As Worksheet: Set wk = ThisWorkbook.Worksheets(sheetNo)
  Dim arrayValue As Collection: Set arrayValue = list

  '�V�[�g���e�擾
  Dim curRowNo As Integer: curRowNo = beginRowNo
  Do While wk.Cells(curRowNo, colNo).Value <> ""
    arrayValue.Add (wk.Cells(curRowNo, colNo).Value)
    curRowNo = curRowNo + 1
  Loop

  Set VarticalValueGetAdd = arrayValue
End Function


' @title maxRowCount
' @param �V�[�g�ԍ�,�J�E���g�J�n�s�ԍ�,��ԍ�
' @sinse 2018.04.27
' @memo �w���1��̃f�[�^������Z�������擾����
'       ��s������Ǝ擾���I������B
Function maxRowCount(sheetNo As Integer, beginRowNo As Long, colNo As Long) as long
  Dim sheet As Worksheet: Set sheet = ThisWorkbook.Worksheets(sheetNo)
  Dim count As Long: count = 0
  Dim curRowNo As Long: curRowNo = beginRowNo
  Do While sheet.Cells(curRowNo, colNo).value <> ""
    count = count + 1
    curRowNo = curRowNo + 1
  Loop
  
  maxRowCount = count
End Function


' �������珬���ɕϊ�
const FRACTION_PATTERN as String = "^(\d+)/(\d+)$"
function calcFractionToDecimal(value as String) as String
	dim converting as String:converting = ""
	' �������ǂ�������
	if isFraction(value) Then
	' �E���ƍ����ŕ����Đ��l�^�ɕϊ�
		dim left as long :left = CLng(getLeft(value))
		dim right as long :right = CLng(getRight(value))
		' ����
		dim calced as Double : calced = left / right
		' ���ʂ𕶎���ɂ��ĕԂ�
		converting = CSTR(calced)
	else
		converting = value
	end if
	calcFractionToDecimal = converting
end function
private function isFraction(value as String) as Boolean
	dim isFraction_ as boolean:isFraction_ = false
	if isMatchRegex(value,FRACTION_PATTERN) Then
		isFraction_=true
	end if 
	isFraction = isFraction_
end function
private function getLeft(value as String) as String
	getLeft = regex(value,FRACTION_PATTERN,"$1")
end function
private function getRight(value as String) as String
	getRight = regex(value,FRACTION_PATTERN,"$2")
end function







' @title FileOut
' @param �o�̓t�@�C����
' @param �t�@�C����, ���e(��s����), �o�͂���p�X
' @sinse 2018.04.27
' @memo �t�@�C���o�͂��܂��B
'       �u�b�N������p�X�� thisWorkbook.path
'       Visual Basic Editor �̃��j���[����m�c�[���n���m�Q�Ɛݒ�n��I�сC
'      �m�Q�Ɖ\�ȃ��C�u�����t�@�C���n�̒����� "Microsoft ActiveX Data Objects x.x Library" �Ƀ`�F�b�N�����܂��B6�n���ǂ�
Sub FileOut(fileName As String, arrayValue As Collection, path As String)
  '�����`�F�b�N
  fileName = FileNameIncrement(fileName, path)

  '�t�@�C������
  Dim datFile As String
  datFile = path & "\" + fileName

  'ADODB.Stream�I�u�W�F�N�g�𐶐�
  Dim adoSt As Object
  Set adoSt = CreateObject("ADODB.Stream")

  '��s��
  Dim strLine As String
  Dim str As Variant

  '�t�@�C���o��
  With adoSt
    .Charset = "UTF-8"
    .LineSeparator = adLF
    .Open
    '�s���[�v
    For Each str In arrayValue
        strLine = ""
        strLine = strLine & str
        .WriteText strLine, adWriteLine
    Next
    
    '---- BOM�Ȃ� ----'
    .Position = 0 '�X�g���[���̈ʒu��0�ɂ���
    .Type = adTypeBinary '�f�[�^�̎�ނ��o�C�i���f�[�^�ɕύX
    .Position = 3 '�X�g���[���̈ʒu��3�ɂ���

    Dim byteData() As Byte '�ꎞ�i�[�p
    byteData = .Read '�X�g���[���̓��e���ꎞ�i�[�p�ϐ��ɕۑ�
    .Close '��U�X�g���[�������i���Z�b�g�j

    .Open '�X�g���[�����J��
    .Write byteData '�X�g���[���Ɉꎞ�i�[�����f�[�^�𗬂�����
    '-----------------'
    
    .SaveToFile datFile
    .Close
  End With

  MsgBox fileName + "�ɏ����o���܂���"
End Sub


' @title autofill
' @param �I�[�g�t�B���̑ΏۂƂȂ�Z��, �I�[�g�t�B���ő��₷�s��
' @sinse 2020.05.20
' @memo �I�[�g�t�B�����s���A�s�̃I�[�g�t�B���͂ł��邪��P�ʂ��Ƃǂ����낤
' @memo �R�s�y�� https://oki2a24.com/2015/12/05/collection-to-array-with-excel-vba/
Sub autofill(sourceRange As Range, size As Long)
  With sourceRange
    .autofill Destination:=.Resize(size)
  End With
End Sub


' @title ClearSheet
' @param �V�[�g�ԍ�
' @sinse 2020.05.27
' @memo �V�[�g�̓��e���폜���܂�
Sub ClearSheet(ByVal sheetNo As Integer)
    With ThisWorkbook.Worksheets(sheetNo)
        .Cells.Clear
    End With
End Sub


' @title deactivateDispUpdating
' @sinse 2020.05.27
' @memo �`��X�V�𖳌���
Sub deactivateDispUpdate()
  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
  End With
End Sub


' @title activateDispUpdating
' @sinse 2020.05.27
' @memo �`��X�V��L��
Sub activateDispUpdate()
  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
  End With
End Sub







' @title ExampleFileOut
' @sinse 2018.10.05
' @memo FileOut�̋L�q��ł��B
Sub ExampleFileOut(outputPath As String)
  ' Application.Calculate
  ' output = ActiveWorkbook.path
  Const FILE_NAME As String = "program_upd.sql"
  Const BEGIN_ROW_NO As Long = 1
  dim sheetNo As Long : sheetNo = CUR_SHEET_NO
  dim sqlColCode as String : sqlColCode = SQL_COL_CODE
  
  Call FileOut(FILE_NAME, VarticalValueGet(sheetNo, BEGIN_ROW_NO, convertColNo(sqlColCode)), outputPath)
End Sub