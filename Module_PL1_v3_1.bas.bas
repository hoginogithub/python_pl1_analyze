' �O���[�o���ϐ�
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

Dim commentStatus   As Integer ' 0:�ʏ�, 1:�I�[�u*/�v���� or �I�[�u"�v���� or �I�[�u'�v����
Dim procStatus      As Integer ' 10:�v���V�W����������, 11:�v���V�W��������, 12:�v���V�W���J�n����, 13: �v���V�W���I�[����
Dim conditionStatus As Integer ' 20:�uIF�v����, 21:�uELSE IF�vor �uELSE�v����
Dim loopStatus      As Integer ' 30:�uDO�v����, 31:�uWHILE�v����, 32:�uTO�v����, 33:�uREPEAT�v����

Dim lineMax         As Long      '�\�[�X�s��
Dim searchWord      As Variant

Type ProcStackType
    Name      As Variant            '�v���V�W����
    Num       As Variant            '�v���V�W���ԍ�
    StartLine As Variant            '�J�n�s��
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
 'F���UT�P�[�XID��t�^
 Dim resultSheet  As Worksheet
 Dim analyzeSheet As Worksheet
 Dim sourceRng    As Range
 Dim startRow     As Long      '�\�[�X�R�[�h��͊J�n �s�i�G�N�Z���s�j
 Dim endRow       As Long      '�\�[�X�R�[�h��͏I�� �s�i�G�N�Z���s�j

    ' ��������
    Set resultSheet = Worksheets("��r����")
    Set analyzeSheet = Worksheets("UT Case ID �̔ԃV�[�g")
    If InitMclSetC <> 0 Then
        Exit Sub
    End If

    'F��}��
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

    '�\�[�X�s��
    startRow = 3
    maxRow = ActiveSheet.Range("E3").End(xlDown).row
    lineMax = maxRow - startRow + 1

    '�\�[�X�R�[�h���
    Call ConvertAnalyzeSheet

    MsgBox ("���s�������܂���")

End Sub

' MclSetC ��������
Function InitMclSetC() As Integer
    ' ���폈��
    InitMclSetC = 0

    '��r���ʃV�[�g�Ŏ��{���Ă��邩�̃`�F�b�N
    InitMclSetC = SheetChk
    If InitMclSetC = -1 Then
        Exit Function
    End If

    '�l�b�k�Z�b�g�̎��s�v�ۊm�F���s���B
    Dim msg, Style, Title, Response
    msg = "�l�b�k�ԍ��̎����t�Ԃ����s���܂��B"
    Style = vbYesNo + vbQuestion + vbDefaultButton2    ' �{�^�����`���܂��B
    Title = "�l�b�k�ԍ��t�Ԏ��s�v��"                ' �^�C�g�����`���܂��B

    ' ���b�Z�[�W��\�����܂��B
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then    ' [�͂�] ���N���b�N���ꂽ�ꍇ�A�������Ȃ��ő��s
    Else    ' [������] ���N���b�N���ꂽ�ꍇ�A�}�N���I��
        InitMclSetC = -1
        Exit Function
    End If

End Function

' UT Case ID �̔ԃV�[�g��ҏW
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
        ' �R�����g���폜
        outData = ""
        tmpData = StrConvSp(Range("UTCaseID�̔ԃe�[�u��[��r����_�ύX��\�[�X_�啶���ϊ�]")(currentLine).Value)
        Call ChangeStatusDelComment(currentLine, tmpData, outData)

        If currentLine Mod 500 = 0 Then
            Application.StatusBar = Trim(Str(currentLine)) & "�s�ڒʉ� (Max: " & Trim(Str(lineMax)) & "�s) (" & Trim(Str((currentLine * 100) / lineMax)) & "%)"
        End If

        tmpData = Trim(outData)
        If tmpData <> "" Then
            ' �v���V�W�� ���
            Call SearchProceduer(currentLine, tmpData)

            ' �������� ���
            Call SearchIfThenElse(currentLine, tmpData)

            ' �J��Ԃ����� ���
            Call SearchLoop(currentLine, tmpData)
        End If
    Next
    Application.StatusBar = False
    Do Until procCnt = 0
        Call DeleteProcStack(lineMax)
    Loop
End Sub

'�R�����g�A�E�g�A��������폜
Private Sub ChangeStatusDelComment(ByVal currentLine As Long, ByRef tmpData As Variant, ByRef outData As Variant)
 Dim result As Boolean
    result = True
    If commentStatus = OP_NORMAL Then
        Call NormallineInDel(tmpData, outData, result)
    Else
        Call SearchDeliminateCode(tmpData, outData, result)
    End If
    Range("UTCaseID�̔ԃe�[�u��[��r����_�ύX��\�[�X_�R�����g������]")(currentLine).Value = outData
End Sub
' ���R�����g�A������폜�������@0:�ʏ�
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
' ���R�����g�A������폜������ ��������
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

' �v���V�W�� ���
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
�@              Call DeleteProcStack(currentLine)
                tmpData = Mid(tmpData, endPos + endLen + 1)
                procStatus = OP_PROC_NAME
            End If
        End If
    Loop
End Sub
' ��1�������ł������� ���� 0 �ł͂Ȃ� (a���ł��������A0�ȏ�Ȃ�True)
Private Function CheckMinPos(a, b, c) As Boolean
    If a > 0 And (a < b Or b = 0) And (a < c Or c = 0) Then
        CheckMinPos = True
    Else
        CheckMinPos = False
    End If
End Function
' �v���V�W���� or ���x���� �擾
Private Sub GetProcName(ByVal currentLine As Long, ByVal namePos As Integer, ByRef tmpData As String)
    ReDim Preserve tmpProc(tmpProcCnt + 1)
    tmpProc(tmpProcCnt).Name       = Trim(Left(tmpData, namePos - 1))
    tmpProc(tmpProcCnt).Num        = procNum + 1
    tmpProc(tmpProcCnt).StartLine  = currentLine
    tmpProcCnt = tmpProcCnt + 1

    tmpData = Trim(Mid(tmpData, namePos + 1))
End Sub
' �v���V�W�������� (���x��������������)
Private Sub SearchProcName(ByVal tmpData,)
' PROC or PROCEDURE ���� (OP_PROC_PROC)
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
' �v���V�W���̖��� ���� (OP_PROC_END)
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
' �v���V�W���E�X�^�b�N��ǉ�
Private Sub SetProcStack(ByVal procPos As Integer, ByVal procLen As Integer, ByRef tmpData As String)
    procCnt = procCnt + 1
    procNum = procNum + 1
    ReDim Preserve procStack(procCnt)
    procStack(procCnt -1) = tmpProc
    tmpData = Trim(Mid(tmpData, procPos + procLen))
End Sub
' �v���V�W���E�X�^�b�N���폜
Private Sub DeleteProcStack(ByVal currentLine As Long)
 Dim row As Long
    For row = procStack(procCnt-1).startLine To currentLine
        if Range("UTCaseID�̔ԃe�[�u��[�v���V�W���ԍ�]").Cells(row,1) = 0 Then
            Range("UTCaseID�̔ԃe�[�u��[�v���V�W����]").Cells(row,1) = procStack(procCnt-1).Name
            Range("UTCaseID�̔ԃe�[�u��[�v���V�W���ԍ�]").Cells(row,1) = procStack(procCnt-1).Num
        End If
    Next
    procCnt = procCnt - 1
    If procCnt = 0 Then
        Erase procStack
    Else
        ReDim Preserve procStack(procCnt)
    End If
End Sub
' �v���V�W���ԍ��A�v���V�W�����Z�b�g
Private Sub SetProcedurData(ByVal row As Long, ByVal endRow As Long)
 Dim idx As Long
    For idx = row To endRow

    Next
End Sub
' �������� ���
Private Sub SearchIfThenElse(ByVal currentLine As Long, ByVal tmpData As String)
End Sub

' �J��Ԃ����� ���
Private Sub SearchLoop(ByVal currentLine As Long, ByVal tmpData As String)
End Sub

Sub Del_CaseId()
Attribute Del_CaseId.VB_ProcData.VB_Invoke_Func = "d\n14"
    'F����폜

    Dim Ret As Integer

    '��r���ʃV�[�g�Ŏ��{���Ă��邩�̃`�F�b�N
    Ret = SheetChk
    If Ret = -1 Then
        Exit Sub
    End If

    Dim msg, Style, Title, Response
    msg = "UT Case ID����폜���܂�"
    Style = vbYesNo + vbQuestion + vbDefaultButton2
    Title = "UT Case ID��폜��"

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
    If SheetName <> "��r����" Then
        MsgBox ("��r���ʃV�[�g�Ŏ��{���Ă�������")
        SheetChk = -1
    End If
End Function


' ���K�\���ɂ��L�[���[�h�̌���
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
' ���K�\�����AVBA(VBScript)�̐��K�\���̃I�u�W�F�N�g�ɕϊ�
'
Private Sub SetREObject(ByVal regPtn As String, ByRef RE As Variant)
   Set RE = CreateObject("VBScript.RegExp")
   With RE
      .Pattern = regPtn          ' �����p�^�[����ݒ�
      .IgnoreCase = False        ' True�̏ꍇ�A�啶���Ə���������ʂ��Ȃ�
      .Global = True             ' ������S�̂�����
   End With

End Sub

'' �e�X�g�c�[��
Public Sub ClearAnalyzeData()
    Range("UTCaseID�̔ԃe�[�u��[��r����_�ύX��\�[�X_�R�����g������]").ClearContents
    Range("UTCaseID�̔ԃe�[�u��[�v���V�W����]").ClearContents
    Range("UTCaseID�̔ԃe�[�u��[�v���V�W���ԍ�]").ClearContents
    Range("UTCaseID�̔ԃe�[�u��[�������򒊏o]").ClearContents
    Range("UTCaseID�̔ԃe�[�u��[�J��Ԃ��������o]").ClearContents
End Sub
