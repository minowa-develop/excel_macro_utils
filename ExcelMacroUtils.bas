'-------------------------------
'-- ExcelMacroUtils ver 0.2.1 --
'-------------------------------

'== Functions ==
'regex(対象文字t, 置換元文字(正規表現使用可能)p, 置換文字r)
'isMatchRegex(対象文字t, 置換元文字(正規表現使用可能)p)
'IsBlankRide(チェックする文字)
'fileNameIncrement(ファイル名(拡張子あり, ドット複数厳禁), チェックするパス)
'getSearchRowNo(検索値, シート番号, カウント開始行, カウント列)
'SWITCH(検索値, 被検索値1, 取得値1, 被検索値2, 取得値2, …, 値がない場合の表示)
'ConvertRange(始点_列番号, 始点_行番号, 終点_列番号, 終点_行番号)
'ConvertColCode(列番号)
'ConvertColNo(列コード)
'VarticalValueGet(シート番号,取得開始行番号,列番号)
'VarticalValueGetAdd(シート番号,取得開始行番号,列番号,追加先collection)
'maxRowCount(シート番号,カウント開始行番号,列番号)
'ExactConmaValues(検索値, 被検索値(コンマ区切り))
'trimRide(文字列)
'getnbsp()
'sdel(文字列)
'getFiles(パス,拡張子)
'ArrayToCollection(変換前(Collection))
'CollectionToArray(変換前(array))
'TEXTJOIN(区切り文字,空白を無視するかどうか,join対象の文字...)

'== sub ==
'FileOut(ファイル名, 内容(一行ごと), 出力するパス)
'autofill(オートフィルの対象となるセル, オートフィルで増やす行数)
'ClearSheet(シート番号)
'deactivateDispUpdate
'activateDispUpdate


Option Explicit

' @title regex
' @param 対象文字t, 置換元文字(正規表現使用可能)p, 置換文字r
' @sinse 2018.04.17
' @memo 正規表現が使える置換関数です。
' @help https://www.sejuku.net/blog/33541
Function regex(target As String, ptn As String, rep As String) as String 
  'RegExpオブジェクトの作成
  Dim REG As Object: Set REG = CreateObject("VBScript.RegExp")

  '正規表現の指定
  With REG
    .PATTERN = ptn         'パターンを指定
    .IgnoreCase = False     '大文字と小文字を区別するか(False)、しないか(True)
    .Global = True          '文字列全体を検索するか(True)、しないか(False)
  End With

  regex = REG.replace(target, rep) '指定した正規表現を第2引数の区切り文字に置換文字
End Function

' @title isMatchRegex
' @param 対象文字t, 置換元文字(正規表現使用可能)p, 置換文字r
' @sinse 2018.04.17
' @memo 正規表現にマッチしたらtrueを返す
Function isMatchRegex(target As String, ptn As String) As Boolean
  'RegExpオブジェクトの作成
  Dim REG As Object: Set REG = CreateObject("VBScript.RegExp")

  '正規表現の指定
  With REG
    .PATTERN = ptn         'パターンを指定
    .IgnoreCase = False     '大文字と小文字を区別するか(False)、しないか(True)
    .Global = True          '文字列全体を検索するか(True)、しないか(False)
  End With

  isMatchRegex = REG.Test(target) '指定した正規表現を第2引数の区切り文字に置換文字
End Function


' @title IsBlankRide
' @param チェックする文字
' @sinse 2018.04.17
' @update 2019.07.08
' @memo 空文字と空白文字の場合trueをかえします。
'       既存のISBLANKは空文字しか見ないなんちゃってぶらんくなのでちゅうい
'       nbsp対応
Function IsBlankRide(str As String) as String 
  '定数
  Dim nbsp As String: nbsp = getnbsp()
  Dim REG As String: REG = "[\s" + nbsp + "]"

  Dim return_flg As Boolean: return_flg = False

  '空白削除
  Dim strRep As String
  strRep = regex(str, REG, "")

  '判定
  If strRep = "" Then return_flg = True

  IsBlankRide = return_flg
End Function


' @title FileNameIncrement
' @param ファイル名(拡張子あり, ドット複数厳禁), チェックするパス
' @sinse 2018.05.10
' @memo 同名のファイルがあった場合(i)形式にする
Function FileNameIncrement(fileName As String, path As String) as String
  '同名チェック
  If Dir(path & "\" + fileName) <> "" Then
    'インクリメント付与
    ' 最後の.をヒットさせるには p:\.([^\.]+)$ でキャプチャ部分が拡張子となる
    ' fileName = regex(fileName, "\.", "(1).")
    fileName = regex(fileName, "\.([^.]+$)", "(1).$1")
    Dim i As Integer: i = 1
    Do While Dir(path & "\" + fileName) <> ""
      'インクリメント
      fileName = regex(fileName, "\([0-9]+\)", "(" + CStr(i) + ")")
      i = i + 1
    Loop
  End If

  FileNameIncrement = fileName
End Function


' @title getSearchRowNo
' @param 検索値,シート番号,カウント開始行,カウント列
' @sinse 2019.01.30
' @memo  検索にヒットした行数(カウント開始行からカウント)を返す。
'        ヒットしない場合は-1を返す
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
' @memo excel2010にswitch関数が使えないので代用。
'       参考コピペ元(https://www.excelspeedup.com/switch/)
' @memo (?<key>[^\t]+)\t(?<value>[^\r\n]+)\r\n → "\k<key>","\k<value>",
' @memo SWITCH(search,key1,value1,key2,value2,…,else_value)
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
' @param 始点_列番号, 始点_行番号, 終点_列番号, 終点_行番号
' @sinse 2019.05.16
' @memo Range形式に変換する 例)A1:B3
Function ConvertRange(colNo_from As Long, rowNo_from As Long, colNo_to As Long, rowNo_to As Long) as String 
  Dim rangeBuf As String: rangeBuf = ConvertColCode(colNo_from) & rowNo_from & ":" & ConvertColCode(colNo_to) & rowNo_to
  ConvertRange = rangeBuf
End Function


' @title ConvertColCode
' @param 列番号
' @sinse 2019.05.16
' @memo 列番号から列コードに変換
Function ConvertColCode(colNo As Long) as String 
  Dim buf As String
  buf = Cells(1, colNo).Address(True, False)
  buf = Left(buf, InStr(buf, "$") - 1)

  ConvertColCode = buf
End Function


' @title ConvertColNo
' @param 列コード
' @sinse 2020.07.15
' @memo 列コードから列番号に変換
Function ConvertColNo(colCode As String) As Long
  ConvertColNo = Range(colCode & "1").Column
End Function


' @title ExactConma
' @param 検索値, 被検索値(コンマ区切り)
' @sinse 2019.05.31
' @memo 被検索値コンマ区切りから検索値をがあるか探す
Function ExactConmaValues(search As String, str As String) as boolean
  Dim ret As Boolean: ret = False
  'conmaconvert
  Dim strList As Variant: strList = Split(str, ",")

  Dim i As Integer: i = 0
  For i = 0 To UBound(strList)
    ret = search = strList(i)
      'tmpがtrueの場合forから出たい
      If ret Then Exit For
  Next i

  ExactConmaValues = ret
End Function


' @title trimRide
' @param 文字列
' @sinse 2019.07.03
' @memo 先頭末尾の\sとnbspを削除する
'       trim関数はnbspに対応していない。
Function trimRide(str As String) as String 
  Dim nbsp As String: nbsp = getnbsp()
  Dim ptn_module As String: ptn_module = "[\s" & nbsp & "]+"
  Dim ptn As String: ptn = "^" & ptn_module & "|" & ptn_module & "$"
  
  trimRide = regex(str, ptn, "")
End Function


' @title getnbsp
' @sinse 2019.07.03
' @memo nbsp取得(no break space 00a0)
Function getnbsp()
  getnbsp = ChrW(160)
End Function


' \s系を取り除く
Function sDel(target As String) As String
  Dim nbsp As String: nbsp = getnbsp()
  Dim ptn As String: ptn = "[\s" & nbsp & vbCrLf & "　]+"
  
  sDel = regex(target, ptn, "")
End Function


' @title getFiles
' @param パス,拡張子
' @sinse 2020.05.20
' @memo 指定したパス内に指定した拡張子のファイルをリストで取得する
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
' @param 配列
' @sinse 2020.05.20
' @memo 配列をcollectionに変換する
' @memo コピペ元 https://oki2a24.com/2016/03/07/array-to-flat-collection-in-excel-vba/
Function ArrayToCollection(ByVal arr As Variant) As Collection
  Set ArrayToCollection = New Collection
  ' 引数が配列でない場合は要素数 1 のコレクションとして返す。
  If Not IsArray(arr) Then
    Call ArrayToCollection.Add(arr)
    Exit Function
  End If

  Dim v As Variant
  For Each v In arr
    If IsArray(v) Then
      ' 配列の要素が配列の場合は再帰的に処理
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
' @memo collectionを配列に変換する
' @memo コピペ元 https://oki2a24.com/2015/12/05/collection-to-array-with-excel-vba/
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
' @memo excel2016で使えないtextjoinを使えるようにする関数
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
' @param シート番号,取得開始行番号,列番号
' @sinse 2018.04.27
' @memo 指定の1列を取得する
'       空行があると取得を終了する。
Function VarticalValueGet(sheetNo As Integer, beginRowNo As Long, colNo As Long) as collection
  Dim wk As Worksheet: Set wk = ThisWorkbook.Worksheets(sheetNo)
  Dim arrayValue As Collection: Set arrayValue = New Collection

  'シート内容取得
  Dim curRowNo As Long: curRowNo = beginRowNo
  Do While wk.Cells(curRowNo, colNo).value <> ""
    arrayValue.Add (wk.Cells(curRowNo, colNo).value)
    curRowNo = curRowNo + 1
  Loop

  Set VarticalValueGet = arrayValue
End Function

' @title VarticalValueGetを既存のcollectionに追加する関数
' @param シート番号,取得開始行番号,列番号,追加先collection
' @sinse 2018.04.27
' @memo 指定の1列を取得する
'       空行があると取得を終了する。
Function VarticalValueGetAdd(sheetNo As Integer, beginRowNo As Long, colNo As Long, list As Collection) as collection
  Dim wk As Worksheet: Set wk = ThisWorkbook.Worksheets(sheetNo)
  Dim arrayValue As Collection: Set arrayValue = list

  'シート内容取得
  Dim curRowNo As Integer: curRowNo = beginRowNo
  Do While wk.Cells(curRowNo, colNo).Value <> ""
    arrayValue.Add (wk.Cells(curRowNo, colNo).Value)
    curRowNo = curRowNo + 1
  Loop

  Set VarticalValueGetAdd = arrayValue
End Function


' @title maxRowCount
' @param シート番号,カウント開始行番号,列番号
' @sinse 2018.04.27
' @memo 指定の1列のデータがあるセル数を取得する
'       空行があると取得を終了する。
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


' 分数から小数に変換
const FRACTION_PATTERN as String = "^(\d+)/(\d+)$"
function calcFractionToDecimal(value as String) as String
	dim converting as String:converting = ""
	' 分数かどうか判定
	if isFraction(value) Then
	' 右側と左側で分けて数値型に変換
		dim left as long :left = CLng(getLeft(value))
		dim right as long :right = CLng(getRight(value))
		' 割る
		dim calced as Double : calced = left / right
		' 結果を文字列にして返す
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
' @param 出力ファイル名
' @param ファイル名, 内容(一行ごと), 出力するパス
' @sinse 2018.04.27
' @memo ファイル出力します。
'       ブックがあるパスは thisWorkbook.path
'       Visual Basic Editor のメニューから［ツール］→［参照設定］を選び，
'      ［参照可能なライブラリファイル］の中から "Microsoft ActiveX Data Objects x.x Library" にチェックを入れます。6系が良い
Sub FileOut(fileName As String, arrayValue As Collection, path As String)
  '同名チェック
  fileName = FileNameIncrement(fileName, path)

  'ファイル生成
  Dim datFile As String
  datFile = path & "\" + fileName

  'ADODB.Streamオブジェクトを生成
  Dim adoSt As Object
  Set adoSt = CreateObject("ADODB.Stream")

  '一行分
  Dim strLine As String
  Dim str As Variant

  'ファイル出力
  With adoSt
    .Charset = "UTF-8"
    .LineSeparator = adLF
    .Open
    '行ループ
    For Each str In arrayValue
        strLine = ""
        strLine = strLine & str
        .WriteText strLine, adWriteLine
    Next
    
    '---- BOMなし ----'
    .Position = 0 'ストリームの位置を0にする
    .Type = adTypeBinary 'データの種類をバイナリデータに変更
    .Position = 3 'ストリームの位置を3にする

    Dim byteData() As Byte '一時格納用
    byteData = .Read 'ストリームの内容を一時格納用変数に保存
    .Close '一旦ストリームを閉じる（リセット）

    .Open 'ストリームを開く
    .Write byteData 'ストリームに一時格納したデータを流し込む
    '-----------------'
    
    .SaveToFile datFile
    .Close
  End With

  MsgBox fileName + "に書き出しました"
End Sub


' @title autofill
' @param オートフィルの対象となるセル, オートフィルで増やす行数
' @sinse 2020.05.20
' @memo オートフィルを行う、行のオートフィルはできるが列単位だとどうだろう
' @memo コピペ元 https://oki2a24.com/2015/12/05/collection-to-array-with-excel-vba/
Sub autofill(sourceRange As Range, size As Long)
  With sourceRange
    .autofill Destination:=.Resize(size)
  End With
End Sub


' @title ClearSheet
' @param シート番号
' @sinse 2020.05.27
' @memo シートの内容を削除します
Sub ClearSheet(ByVal sheetNo As Integer)
    With ThisWorkbook.Worksheets(sheetNo)
        .Cells.Clear
    End With
End Sub


' @title deactivateDispUpdating
' @sinse 2020.05.27
' @memo 描画更新を無効化
Sub deactivateDispUpdate()
  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
  End With
End Sub


' @title activateDispUpdating
' @sinse 2020.05.27
' @memo 描画更新を有効
Sub activateDispUpdate()
  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
  End With
End Sub







' @title ExampleFileOut
' @sinse 2018.10.05
' @memo FileOutの記述例です。
Sub ExampleFileOut(outputPath As String)
  ' Application.Calculate
  ' output = ActiveWorkbook.path
  Const FILE_NAME As String = "program_upd.sql"
  Const BEGIN_ROW_NO As Long = 1
  dim sheetNo As Long : sheetNo = CUR_SHEET_NO
  dim sqlColCode as String : sqlColCode = SQL_COL_CODE
  
  Call FileOut(FILE_NAME, VarticalValueGet(sheetNo, BEGIN_ROW_NO, convertColNo(sqlColCode)), outputPath)
End Sub