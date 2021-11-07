' グローバル変数
Const OP_NORMAL        As Integer = 0
Const OP_DELIMINATER   As Integer = 1

Const OP_PROC_INIT     As Integer = 10
Const OP_PROC_NAME     As Integer = 11
Const OP_PROC_PROC     As Integer = 12
Const OP_PROC_END      As Integer = 13

Const OP_COND_IF       As Integer = 20
Const OP_COND_ELSE     As Integer = 21

Const OP_LOOP_DO       As Integer = 30
Const OP_LOOP_WHILE    As Integer = 31
Const OP_LOOP_TO       As Integer = 32
Const OP_LOOP_REPEAT   As Integer = 33

Dim commentStatus   As Integer ' 0:通常, 1:終端「*/」検索 or 終端「"」検索 or 終端「'」検索
Dim procStatus      As Integer ' 10:プロシジャ検索初期, 11:プロシジャ名検索, 12:プロシジャ開始検索, 13: プロシジャ終端検索
Dim conditionStatus As Integer ' 20:「IF」検索, 21:「ELSE IF」or 「ELSE」検索
Dim loopStatus      As Integer ' 30:「DO」検索, 31:「WHILE」検索, 32:「TO」検索, 33:「REPEAT」検索

Dim lineMax         As Long      'ソース行数
Dim searchWord      As Variant

Type ProcStackType
    Name      As Variant            'プロシジャ名
    Num       As Variant            'プロシジャ番号
    StartLine As Variant            '開始行数
End Type
Dim procStack()   As ProcStackType
Dim procCnt       As Integer
Dim procNum       As Integer

Dim tmpProc()     As ProcStackType
Dim tmpProcCnt    As Integer

Dim ifStack()     As String
Dim ifPos         As Integer

Public Sub MclSetC()
Attribute MclSetC.VB_ProcData.VB_Invoke_Func = "q\n14"
 'F列にUTケースIDを付与
 Dim resultSheet  As Worksheet
 Dim analyzeSheet As Worksheet
 Dim sourceRng    As Range
 Dim startRow     As Long      'ソースコード解析開始 行（エクセル行）
 Dim endRow       As Long      'ソースコード解析終了 行（エクセル行）

    ' 初期処理
    Set resultSheet = Worksheets("比較結果")
    Set analyzeSheet = Worksheets("UT Case ID 採番シート")
    If InitMclSetC <> 0 Then
        Exit Sub
    End If

    'F列挿入
    If Cells(2, 6) <> "UT Case ID" Then
        Columns(6).Insert
        Cells(2, 6) = "UT Case ID"
        Cells(2, 6).HorizontalAlignment = xlCenter
        Cells(2, 6).VerticalAlignment = xlCenter
        Cells(2, 6).Interior.ColorIndex = 43
        Columns("F").ColumnWidth = 20
        Columns("F").Font.ColorIndex = 1
        Columns("F").Font.Size = 11
    End If

    'ソース行数
    startRow = 3
    maxRow = ActiveSheet.Range("E3").End(xlDown).row
    lineMax = maxRow - startRow + 1

    'ソースコード解析
    Call ConvertAnalyzeSheet

    MsgBox ("実行完了しました")

End Sub

' MclSetC 初期処理
Function InitMclSetC() As Integer
    ' 正常処理
    InitMclSetC = 0

    '比較結果シートで実施しているかのチェック
    InitMclSetC = SheetChk
    If InitMclSetC = -1 Then
        Exit Function
    End If

    'ＭＣＬセットの実行要否確認を行う。
    Dim msg, Style, Title, Response
    msg = "ＭＣＬ番号の自動付番を実行します。"
    Style = vbYesNo + vbQuestion + vbDefaultButton2    ' ボタンを定義します。
    Title = "ＭＣＬ番号付番実行要否"                ' タイトルを定義します。

    ' メッセージを表示します。
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then    ' [はい] がクリックされた場合、何もしないで続行
    Else    ' [いいえ] がクリックされた場合、マクロ終了
        InitMclSetC = -1
        Exit Function
    End If

End Function

' UT Case ID 採番シートを編集
Private Sub ConvertAnalyzeSheet()
 Dim tmpData           As String
 Dim outData           As String
 Dim currentLine       As Long
 Dim idx               As Integer

    commentStatus   = OP_NORMAL
    procStatus      = OP_PROC_INIT
    conditionStatus = OP_COND_IF
    loopStatus      = OP_LOOP_DO

    procNum = 0
    procCnt = 0
    tmpProcCnt = 0

    For currentLine = 1 To lineMax
        ' コメント文削除
        outData = ""
        tmpData = StrConvSp(Range("UTCaseID採番テーブル[比較結果_変更後ソース_大文字変換]")(currentLine).Value)
        Call ChangeStatusDelComment(currentLine, tmpData, outData)

        If currentLine Mod 500 = 0 Then
            Application.StatusBar = Trim(Str(currentLine)) & "行目通過 (Max: " & Trim(Str(lineMax)) & "行) (" & Trim(Str((currentLine * 100) / lineMax)) & "%)"
        End If

        tmpData = Trim(outData)
        If tmpData <> "" Then
            ' プロシジャ 解析
            Call SearchProceduer(currentLine, tmpData)

            ' 条件分岐 解析
            Call SearchIfThenElse(currentLine, tmpData)

            ' 繰り返し処理 解析
            Call SearchLoop(currentLine, tmpData)
        End If
    Next
    Application.StatusBar = False
    Do Until procCnt = 0
        Call DeleteProcStack(lineMax)
    Loop
End Sub

'コメントアウト、文字列を削除
Private Sub ChangeStatusDelComment(ByVal currentLine As Long, ByRef tmpData As Variant, ByRef outData As Variant)
 Dim result As Boolean
    result = True
    If commentStatus = OP_NORMAL Then
        Call NormallineInDel(tmpData, outData, result)
    Else
        Call SearchDeliminateCode(tmpData, outData, result)
    End If
    Range("UTCaseID採番テーブル[比較結果_変更後ソース_コメント文除去]")(currentLine).Value = outData
End Sub
' ＜コメント、文字列削除処理＞　0:通常
Sub NormallineInDel(ByRef tmpData As Variant, ByRef outData As Variant, ByRef result As Boolean)
 Dim keywords As Variant: keywords = Array("/*", """", "'")
 Dim deliminates As Variant: deliminates = Array("*/", """", "'")
 Dim keyword As Variant
 Dim keyCnt As Integer
 Dim pos As Integer
    result = True
    commentStatus = OP_NORMAL
    If Len(Trim(tmpData)) = 0 Then
        result = False
        Exit Sub
    End If
    keyCnt = 0
    For Each keyword In keywords
        pos = Instr(tmpData, keyword)
        If  pos = 0 Then
        Else
            If pos <= Len(keyword) Then
            Else
                outData = outData & Left(tmpData, pos - Len(keyword)) + " "
            End If
            searchWord = deliminates(keyCnt)
            tmpData = Mid(tmpData, pos + Len(keyword))
            Call SearchDeliminateCode(tmpData, outData, result)
            If result = False Then
                Exit Sub
            End If
        End If
        keyCnt = keyCnt + 1
    Next keyword
    outData = outData & tmpData
    tmpData = ""
End Sub
' ＜コメント、文字列削除処理＞ 末尾検索
Private Sub SearchDeliminateCode(ByRef tmpData As Variant, ByRef outData As Variant, ByRef result As Boolean)
 Dim pos As Integer
 Dim searchWordLen As Integer
    result = True
    commentStatus = OP_DELIMINATER
    pos = InStr(tmpData, searchWord)
    If pos = 0 Then
         tmpData = ""
         result = False
         Exit Sub
    End If

    commentStatus = OP_NORMAL
    searchWordLen = Len(searchWord)
    searchWord = ""
    If Len(tmpData) = searchWordLen Then
        tmpData = ""
        result = False
        Exit Sub
    End If
    tmpData = Mid(tmpData, pos + searchWordLen)
    Call NormallineInDel(tmpData, outData, result)
End Sub

' プロシジャ 解析
Private Sub SearchProceduer(ByVal currentLine As Long, ByVal tmpData As String)
 Dim pos      As Integer
 Dim namePos  As Integer: namePos = -1
 Dim procPos  As Integer: procPos = -1
 Dim endPos   As Integer: endPos  = -1
 Dim procLen  As Integer: procLen = 0
 Dim endLen   As Integer: endLen  = 0

    Do Until namePos = 0 And procPos = 0 And endPos = 0
        If procCnt = 0 And procStatus = OP_PROC_INIT Then
            namePos = InStr(tmpData, ":")
            If namePos <> 0 Then
               Call GetProcName(currentLine, namePos, tmpData)
               procStatus = OP_PROC_PROC
            Else
               Exit Sub
            End If
        Else
            namePos = SearchProcName(tmpData)
            procPos = SearchProcProc(tmpData, procLen)
            endPos  = SearchProcEnd(tmpData, endLen)
            If namePos = 0 And procPos = 0 And endPos = 0 Then
                Exit Sub
            End If
            If CheckMinPos(namePos, procPos, endPos) Then
                Call GetProcName(currentLine, namePos, tmpData)
                procStatus = OP_PROC_PROC
            ElseIf CheckMinPos(procPos, namePos, endPos) Then
                Call SetProcStack(procPos, procLen, tmpData)
                procStatus = OP_PROC_END
            Else
　              Call DeleteProcStack(currentLine)
                tmpData = Mid(tmpData, endPos + endLen + 1)
                procStatus = OP_PROC_NAME
            End If
        End If
    Loop
End Sub
' 第1引数が最も小さい かつ 0 ではない (aが最も小さく、0以上ならTrue)
Private Function CheckMinPos(a, b, c) As Boolean
    If a > 0 And (a < b Or b = 0) And (a < c Or c = 0) Then
        CheckMinPos = True
    Else
        CheckMinPos = False
    End If
End Function
' プロシジャ名 or ラベル名 取得
Private Sub GetProcName(ByVal currentLine As Long, ByVal namePos As Integer, ByRef tmpData As String)
    ReDim Preserve tmpProc(tmpProcCnt + 1)
    tmpProc(tmpProcCnt).Name       = Trim(Left(tmpData, namePos - 1))
    tmpProc(tmpProcCnt).Num        = procNum + 1
    tmpProc(tmpProcCnt).StartLine  = currentLine
    tmpProcCnt = tmpProcCnt + 1

    tmpData = Trim(Mid(tmpData, namePos + 1))
End Sub
' プロシジャ名検索 (ラベルも引っかかる)
Private Sub SearchProcName(ByVal tmpData,)
' PROC or PROCEDURE 検索 (OP_PROC_PROC)
Private Function SearchProcProc(ByVal tmpData As String, ByRef matchLen As Integer) As Integer
 Dim wkPtn As String
    SearchProcProc = 0
    matchLen = 0
    If procStatus <> OP_PROC_PROC Then
        Exit Function
    End If
    wkPtn = "(PROCEDURE\s?;)|(PROCEDURE\s+)|(PROCEDURE\s?$)|(PROC\s?;)|(PROC\s+)|(PROC\s?$)"
    Call CheckSearchREKeyWord(wkPtn, tmpData, SearchProcProc, matchLen)
End Function
' プロシジャの末尾 検索 (OP_PROC_END)
Private Function SearchProcEnd(ByVal tmpData As String, ByRef matchLen As Integer) As Integer
 Dim wkPtn As String
    SearchProcEnd = 0
    matchLen = 0
    If procCnt = 0 Then
        Exit Function
    End If
    wkPtn = "END\s+" & procStack(procCnt-1).Name & "\s?;"
    Call CheckSearchREKeyWord(wkPtn, tmpData, SearchProcEnd, matchLen)
End Function
' プロシジャ・スタックを追加
Private Sub SetProcStack(ByVal procPos As Integer, ByVal procLen As Integer, ByRef tmpData As String)
    procCnt = procCnt + 1
    procNum = procNum + 1
    ReDim Preserve procStack(procCnt)
    procStack(procCnt -1) = tmpProc
    tmpData = Trim(Mid(tmpData, procPos + procLen))
End Sub
' プロシジャ・スタックを削除
Private Sub DeleteProcStack(ByVal currentLine As Long)
 Dim row As Long
    For row = procStack(procCnt-1).startLine To currentLine
        if Range("UTCaseID採番テーブル[プロシジャ番号]").Cells(row,1) = 0 Then
            Range("UTCaseID採番テーブル[プロシジャ名]").Cells(row,1) = procStack(procCnt-1).Name
            Range("UTCaseID採番テーブル[プロシジャ番号]").Cells(row,1) = procStack(procCnt-1).Num
        End If
    Next
    procCnt = procCnt - 1
    If procCnt = 0 Then
        Erase procStack
    Else
        ReDim Preserve procStack(procCnt)
    End If
End Sub
' プロシジャ番号、プロシジャ名セット
Private Sub SetProcedurData(ByVal row As Long, ByVal endRow As Long)
 Dim idx As Long
    For idx = row To endRow

    Next
End Sub
' 条件分岐 解析
Private Sub SearchIfThenElse(ByVal currentLine As Long, ByVal tmpData As String)
End Sub

' 繰り返し処理 解析
Private Sub SearchLoop(ByVal currentLine As Long, ByVal tmpData As String)
End Sub

Sub Del_CaseId()
Attribute Del_CaseId.VB_ProcData.VB_Invoke_Func = "d\n14"
    'F列を削除

    Dim Ret As Integer

    '比較結果シートで実施しているかのチェック
    Ret = SheetChk
    If Ret = -1 Then
        Exit Sub
    End If

    Dim msg, Style, Title, Response
    msg = "UT Case ID列を削除します"
    Style = vbYesNo + vbQuestion + vbDefaultButton2
    Title = "UT Case ID列削除可否"

    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then
        If Cells(2, 6) = "UT Case ID" Then
            Columns(6).Delete
        End If
    End If

End Sub

Private Function StrConvSp(ByVal strData As String) As String
    Dim strChar As String
    Dim sQuestion As String
    Dim i As Long

    StrConvSp = ""
    sQuestion = Chr(63) '?
    If strData = "" Then
        Exit Function
    End If

    For i = 1 To Len(strData)
        strChar = Mid(strData, i, 1)
        If strChar <> sQuestion And _
           Asc(strChar) = Asc(sQuestion) Then
            StrConvSp = StrConvSp + " "
        Else
            StrConvSp = StrConvSp + strChar
        End If

    Next

End Function

Private Function SheetChk() As Integer
    Dim SheetName As String

    SheetChk = 0
    SheetName = ActiveSheet.Name
    If SheetName <> "比較結果" Then
        MsgBox ("比較結果シートで実施してください")
        SheetChk = -1
    End If
End Function


' 正規表現によるキーワードの検索
Private Sub CheckSearchREKeyWord(ByVal rePtn As String, ByVal tmpData As String, ByRef pos As Integer, ByRef matchLen As Integer)
 Dim RE    As Variant
 Dim mc    As Variant
 Dim m     As Variant
 Dim i     As Integer

    pos = 0
    matchLen = 0
    mStr = ""
    Call SetREObject(rePtn, RE)
    Set mc = RE.Execute(tmpData)
    If mc.Count = 1 Then
        For Each m In mc
            pos = InStr(tmpData, m.Value)
            matchLen = Len(m.Value)
        Next
    End If
End Sub
'
' 正規表現を、VBA(VBScript)の正規表現のオブジェクトに変換
'
Private Sub SetREObject(ByVal regPtn As String, ByRef RE As Variant)
   Set RE = CreateObject("VBScript.RegExp")
   With RE
      .Pattern = regPtn          ' 検索パターンを設定
      .IgnoreCase = False        ' Trueの場合、大文字と小文字を区別しない
      .Global = True             ' 文字列全体を検索
   End With

End Sub

'' テストツール
Public Sub ClearAnalyzeData()
    Range("UTCaseID採番テーブル[比較結果_変更後ソース_コメント文除去]").ClearContents
    Range("UTCaseID採番テーブル[プロシジャ名]").ClearContents
    Range("UTCaseID採番テーブル[プロシジャ番号]").ClearContents
    Range("UTCaseID採番テーブル[条件分岐抽出]").ClearContents
    Range("UTCaseID採番テーブル[繰り返し処理抽出]").ClearContents
End Sub
