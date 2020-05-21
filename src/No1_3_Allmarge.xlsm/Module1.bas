Attribute VB_Name = "Module1"

Dim LogCount As Long
Option Explicit

Sub SplitFiles()

    Dim FN1 As String          '���̃u�b�N�̃t�@�C����
    Dim MacroWS As Worksheet   '���̃u�b�N�̃V�[�g
    Dim Wb_new As Workbook     '������̃u�b�N
    Dim WS As Worksheet        '�������f�[�^�̃V�[�g
    Dim rowsData As Long       '�������u�b�N�̃f�[�^���i�s���j
    Dim colsData As Long       '�������u�b�N�̃f�[�^���i�񐔁j
    Dim R_Data2 As Long        '�������u�b�N�̃f�[�^�J�n�s�i���f�[�^�J�n�s�j
    Dim Ko As Long             '�����t�@�C�����i�W�̐��j

    Dim Wb_Data As Workbook    '1. �������u�b�N
    Dim R_Data1 As Long        '2. �������u�b�N�̃f�[�^�J�n�s�i�^�C�g���s�j
    Dim Path As String         '3. �����f�[�^�ۑ���
    Dim C_Group As String      '4. �O���[�v�Ώۗ�
    Dim My_Group As String     '5. ���W��
    Dim Uni_Word As String     '6. ���j�[�N���[�h
    Dim FN2 As String          '7. ������u�b�N�̃t�@�C����
    Dim PSW As String          '8. �ǂݎ��p�X���[�h
    
    '�l���Z�b�g
    FN1 = ActiveWorkbook.Name
    Set MacroWS = Workbooks(FN1).Worksheets("sheet1")
    Set Wb_Data = Workbooks(MacroWS.Range("C3").Value)
    Set WS = Wb_Data.Worksheets(1)
    R_Data1 = MacroWS.Range("C4")
    R_Data2 = MacroWS.Range("C4") + 1
    Path = MacroWS.Range("C5")
    C_Group = MacroWS.Range("C6")
    My_Group = MacroWS.Range("C7")
    Uni_Word = MacroWS.Range("C8")
    FN2 = MacroWS.Range("C9")
    PSW = MacroWS.Range("C10")
    
    Application.ScreenUpdating = False  '��ʂ��Œ肵�č�����
    
    '�ŏI�s�A�ŏI��̎擾
    Wb_Data.Activate
    rowsData = WS.Cells(Rows.Count, 1).End(xlUp).Row
    colsData = WS.Cells(R_Data1, Columns.Count).End(xlToLeft).Column
    
    '�W���Ń\�[�g
    WS.Range(Rows(R_Data1), Rows(rowsData)).Sort _
        Key1:=Range(C_Group & R_Data1), _
        Order1:=xlAscending, _
        Header:=xlYes, _
        Orientation:=xlTopToBottom
    
    '�W�����ƂɃt�@�C���𕪊����ۑ�
    Do
        '���t�@�C���̃f�[�^�J�n�s�i���ڍs�j���R�s�[���A�V�K�G�N�Z���u�b�N�ɓ\��t��
        Wb_Data.Activate
        WS.Range(Cells(R_Data1, 1), Cells(R_Data1, colsData)).Copy
        Workbooks.Add
        ActiveSheet.Paste Range("A2") '2�s�ڈȍ~�Ƀf�[�^���L�ځi1�s�ڂ̓��j�[�N���[�h�L�ڗp�ɋ󂯂Ă����j
        Set Wb_new = ActiveWorkbook
        
        '�P�W���̂ݒ��o���A�t�@�C������ݒ肵�ĕۑ�
        Wb_Data.Activate
        Ko = WorksheetFunction.CountIf(Columns(C_Group), Cells(R_Data2, C_Group)) '�P�W���̃f�[�^�����Z�o
        Range(Cells(R_Data2, "A"), Cells(R_Data2 + Ko - 1, colsData)).Copy        '�P�W���̃f�[�^�����R�s�[
        Wb_new.Activate
        ActiveSheet.Paste Range("A3")                                             '�V�K�u�b�N��3�s�ڈȉ��ɒl�\��t��
        If Cells(3, C_Group) = My_Group Then
            Range("A1").Value = Uni_Word                                          '���W�̃t�@�C���̂݁AA1�Z���Ƀ��j�[�N���[�h���L��
        End If
        Wb_new.SaveAs Filename:=Path & Cells(3, C_Group) & FN2 & ".xlsx", _
        Password:=PSW                                                             '�w�肵���t�H���_�[�ɕۑ�
        
        Wb_new.Close

        R_Data2 = R_Data2 + Ko
        Loop While Cells(R_Data2, C_Group) <> ""
    
    MsgBox "������������"

    Application.ScreenUpdating = True

End Sub


'---------------------------------
'�쐬���F2020/04/30
'�쐬�ҁF���c
'���g�p����
'�^�u���j���[�́u�c�[���v�|�u�Q�Ɛݒ�v�ŁuMicrosoft Scripting Runtime�v�𓱓��i�`�F�b�N�j���邱�ƁB
'���O�������
'�E�����L�[���[�h�͑Ώۃt�@�C����A1�s�Ɋi�[����Ă�����̂Ƃ���B
'�E�����Ώۃt�@�C���̓G�N�Z���t�@�C���i.xls/.xls*)�݂̂Ƃ���B
'�E�����Ώۂ͎w�肵���t�H���_�������i�T�u�t�H���_�͊܂܂Ȃ��j�̑S�G�N�Z���t�@�C���Ƃ���B
'���C���v�b�g�E�A�E�g�v�b�g��
'�C���v�b�g�@�F�����i�����Ώۃt�H���_�����͖{�t�@���N�V�������Ŏ擾����j
'�A�E�g�v�b�g�F�w��L�[���[�h���܂ރt�@�C�����i��΃p�X�Ŗ{���[�N�u�b�N����"Result"�V�[�g�֏������ށj
'���A�b�v�f�[�g��
'5/18�A�b�v�f�[�g�F������Ή��A�T�u�t�H���_�Ή���
'---------------------------------


Sub SearchFiles()
 Dim ResultTab_Row As Long 'Result�V�[�g�̏����Ώۍs�p�ϐ��i�f�t�H���g�l�̓t�H�[�}�b�g�ɍ��킹�āu2�v�Ŏw�肷��j
 Dim LogTab_Row As Long 'Log�V�[�g�̏����Ώۍs�p�ϐ��i�f�t�H���g�l�̓t�H�[�}�b�g�ɍ��킹�āu2�v�Ŏw�肷��j
 Dim SearchWord As String '�����L�[���[�h�p�ϐ��i�C���v�b�g�{�b�N�X�Ŏ擾����j
 Dim SearchFolder As String '�����Ώۂ̃t�H���_���p�ϐ��i�_�C�����O�Ŏ擾�j
 Dim Get_Foldername As String '�t�@�C���_�C�����O�ɂĎ擾�����Ώۃt�H���_�p�X
 Dim Get_Foldername2 As String '�t�H���_�p�X���H�p�ϐ�
 Dim StrLen As Long '�����񒷌v�Z�p�ϐ�
 Dim Get_Filename As String '�擾�����t�@�C�������ꎞ�I�Ɋi�[����ϐ�
 Dim LoopCount_1 'Loop�J�E���g�p�ϐ�
 Dim LoopCount_2 'Loop�J�E���g�p�ϐ�
 Dim Check_sheet As Worksheet '���[�N�V�[�g�����p�ϐ�
 Dim SearchFilename As String '�����ΏۂƂ��ăt�@�C���I�[�v������ۂɃt�@�C�������i�[����ϐ�
 
 ResultTab_Row = 3 '�����l���R�Őݒ�
 LogTab_Row = 1  '�����l���P�Őݒ�
 LogCount = 1
 
 
 '�����Ώۂ̃t�H���_���w��
 MsgBox "�����Ώۃt�H���_��I�����Ă��������B�I�������t�H���_�z���ɂ���T�u�t�H���_�������ΏۂƂȂ�܂��B"
 With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
        Get_Foldername = .SelectedItems(1)
    Else
        MsgBox "�I�����܂��B�Ď��s���Ă��������B"
        Exit Sub
    End If
 End With

 '������ł̓���ۏ�
 If Left(Get_Foldername, 11) = "http://prdo" Then
    StrLen = Len(Get_Foldername)
    Get_Foldername2 = Right(Get_Foldername, StrLen - 28)
    SearchFolder = "G:" & Get_Foldername2
 Else
    SearchFolder = Get_Foldername
 End If
 
 SearchFolder = SearchFolder & "\"
 
 '�����L�[���[�h���C���v�b�g�{�b�N�X�Ŏ擾
 Do While SearchWord = "" '�L�[���[�h��1�����ȏ���͂����܂Ń��[�v
    SearchWord = InputBox("��������L�[���[�h����͂��Ă�������", "�L�[���[�h����", "") '�C���v�b�g�{�b�N�X�Ŏ擾
    If SearchWord = "" Then  '�����L�[���[�h����͂���Ȃ������ꍇ
        MsgBox "�L�[���[�h��1�����ȏ���͂��Ă��������B�Ȃ��A�L�����Z���͂ł��܂���B" '�G���[���b�Z�[�W��\������
    End If
 Loop
 
 '��ʍX�V�E��ʌx�����I�t�ɂ���
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 
 '�擾�����L�[���[�h��Result�V�[�g�ɏ�������
 ThisWorkbook.Worksheets("Result").Range("B1") = SearchWord

 'Result/Log�V�[�g���N���A����
 ThisWorkbook.Worksheets("Result").Range("B3:C102").ClearContents
 ThisWorkbook.Worksheets("Log").Columns("A").ClearContents

'5/18 �T�u�t�H���_�����ɔ����p�~�iSearchFolder�֐��ɂđΉ��j
' �w�肵���t�H���_�����ɂ���G�N�Z���t�@�C�������擾���Log�V�[�g�֏����o��
' Get_Filename = Dir(SearchFolder & "*.xls") '�w��t�H���_�z����.xls���܂ރt�@�C�����擾�i1�t�@�C�����j
' LoopCount_1 = LogTab_Row '���[�v�p�ϐ��Ƀ��O�V�[�g�����l���Z�b�g
'
' Do While Get_Filename <> "" '�w�肵���t�H���_������.xls���܂ރt�@�C������S�Ď擾����܂Ń��[�v
'    If Get_Filename <> ThisWorkbook.Name Then '�{�}�N���t�@�C���i���ꖼ�̃t�@�C���j�͌����ΏۊO�Ƃ���B
'        ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) = SearchFolder & Get_Filename '�擾�����t�@�C������Log�V�[�g�֏�������
'        LoopCount_1 = LoopCount_1 + 1 '���s�֐i��
'    End If
'    Get_Filename = Dir() '�c��̃t�@�C�������擾�i1�t�@�C�����j
' Loop
 
 Call FolderSearch(SearchFolder)
 
 'Log�V�[�g�ɏ������܂ꂽ�t�@�C����S�ĊJ���ă`�F�b�N����
 LoopCount_1 = LogTab_Row
 LoopCount_2 = ResultTab_Row 'Result�V�[�g�̏����l���Z�b�g
 Do While ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) <> ""  'Log�V�[�g�ɋL�ڂ��ꂽ�t�@�C���ւ̏������S�ďI���܂Ń��[�v
    Workbooks.Open ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) '�G�N�Z���t�@�C�����J���iLog�V�[�g�̏ォ�珇�j
    
    '�e�V�[�g��A1�Z���ɃL�[���[�h���܂܂�Ă��邩�`�F�b�N���A�܂܂�Ă���΃t�@�C������Result�V�[�g�ɏ�������
    For Each Check_sheet In ActiveWorkbook.Worksheets '�S�V�[�g�ւ̏�������������܂Ń��[�v
        Check_sheet.Activate 'Sheet���A�N�e�B�x�[�V��������i�S�V�[�g�����̂��߂ɕK�v�ȏ����j
        If ActiveSheet.Range("A1") = SearchWord Then 'A1�Z���ɋL�ڂ��ꂽ�L�[���[�h�Ɠ��͂����������[�h����v�����ꍇ
            ThisWorkbook.Worksheets("Result").Range("B" & LoopCount_2) = ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) 'Result�V�[�g�Ƀt�@�C�����i��΃p�X�j�����
            ThisWorkbook.Worksheets("Result").Range("C" & LoopCount_2) = ActiveSheet.Name 'Result�V�[�g�Ƀt�@�C�����i��΃p�X�j�����
            LoopCount_2 = LoopCount_2 + 1 'Result�V�[�g�̓��͍s��i�߂�
            Exit For '�����ꂩ�P�̃V�[�g�ŃL�[���[�h�����v�����ꍇ�͓��Y�t�@�C���ւ̌����������I������B
        End If
    Next Check_sheet '���̃V�[�g�ֈړ�����
    
    ActiveWorkbook.Close '�������I������t�@�C�������
    LoopCount_1 = LoopCount_1 + 1 'Log�V�[�g�̍s����i�߂�
 Loop
 
 '�\��t���Ώۃt�@�C�������݂��Ȃ������ꍇ�́A�\��t�������O�ɏI��������B
 If LoopCount_2 = 3 Then
    Application.ScreenUpdating = True
    MsgBox "�w�肳�ꂽ�L�[���[�h�ɍ��v����t�@�C���͑��݂��܂���ł����B�������I�����܂��B"
    Exit Sub
 End If
 
 Call FileAggregation '�t�@�C���������ݏ��������s
 
 '��ʍX�V�E��ʌx�����I���ɂ���
 Application.ScreenUpdating = True
 Application.DisplayAlerts = True
 
 '�������b�Z�[�W��\������
 MsgBox "�������������܂���"
 
End Sub

Public Sub FolderSearch(TargetDir As String) '�T�u�t�H���_�[���܂ޑΏۃt�@�C���i�G�N�Z���t�@�C���j�ꗗ�擾
 Dim FSO As Object
 Dim Folder As Object
 Dim SubFolder As Object
 Dim Filename As Object
 Dim TmpName As String
 Dim CheckName As String
 
 
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Folder = FSO.GetFolder(TargetDir)
 
 For Each SubFolder In Folder.SubFolders
    Call FolderSearch(SubFolder.Path)
 Next SubFolder
 
 For Each Filename In Folder.Files
    TmpName = Filename
    CheckName = Mid(TmpName, InStrRev(TmpName, "\") + 1)
    
    If CheckName <> ThisWorkbook.Name Then '�{�}�N���t�@�C���i���ꖼ�̃t�@�C���j�͌����ΏۊO�Ƃ���B
        If TmpName Like "*.xls*" Then
            ThisWorkbook.Worksheets("Log").Range("A" & LogCount) = TmpName '�擾�����t�@�C������Log�V�[�g�֏�������
            LogCount = LogCount + 1 '���s�֐i��
        End If
    End If
 Next Filename

End Sub



Sub FileAggregation()

    With Application
        .ScreenUpdating = False '��ʍX�V����
        .EnableEvents = False   '�C�x���g�}�~
        .Calculation = xlCalculationManual  '�v�Z�蓮��
    End With
 
    Dim Search_sname As String  '�������ʂ��L�^����Ă���V�[�g��
    Search_sname = "result"

    Dim resultnum As Long   '�������ʐ�
    Dim a As Long   '���[�v�p�ϐ�
    Dim FirstRow As Long  '�\�t�����V�[�g�̃R�s�[�͈͏����s�����i�[
    Dim NextRow As Long   '�W��V�[�g�ւ̓\�t����̍s�����i�[
    Dim Filepath() As String      '�iPJ������, �V�[�g���j���i�[
    Dim LastRow As Long '�\�t�����̗\�Z�V�[�g�ŏI�s���i�[
    Dim LastCol As Long '�\�t�����̗\�Z�V�[�g�ŏI����i�[


    '�������ʂ̃t�@�C���p�X��FilePath�z��֊i�[
    With ThisWorkbook.Worksheets(Search_sname)
        resultnum = 0
    
        '�������ʐ��̊m�F
        Do Until .Cells(resultnum + 3, 2) = ""
            resultnum = resultnum + 1
        Loop
        
        '�������ʂ�0���̏ꍇ�����I��
        If resultnum = 0 Then
            MsgBox "�������ʂ�0���̂��ߏ������I�����܂�"
            End
        End If
         
        resultnum = resultnum - 1
        ReDim Filepath(resultnum, 1) '�z��̗v�f����ݒ�
        
        For a = 0 To resultnum
            Filepath(a, 0) = .Cells(a + 3, 2) '�t�@�C���p�X�̊i�[
            Filepath(a, 1) = .Cells(a + 3, 3) '�V�[�g���̊i�[
        Next a
    End With
            
    FirstRow = 3  '�\�t�����V�[�g�̃R�s�[�͈͏����s����ݒ�
    NextRow = 2 '�W��V�[�g�̓\�t���揉���s����ݒ�
    
    '�W��V�[�g�փR�s�[
    For a = 0 To resultnum
        With Workbooks.Open(Filepath(a, 0))
            '�\��t�����͈̔͂��擾
            LastRow = .Worksheets(Filepath(a, 1)).Cells(Rows.Count, 1).End(xlUp).Row
            LastCol = .Worksheets(Filepath(a, 1)).Cells(LastRow, Columns.Count).End(xlToLeft).Column
            '�\��t�����̃R�s�[
            .Worksheets(Filepath(a, 1)).Range(Worksheets(Filepath(a, 1)).Cells(FirstRow, 1), _
                                              Worksheets(Filepath(a, 1)).Cells(LastRow, LastCol)).Copy
            '���̃u�b�N�̏W��Z���֒l�\��t��
            ThisWorkbook.Worksheets("�W��").Cells(NextRow, 1).PasteSpecial xlPasteAll
            '����\�t���ʒu�̐ݒ�
            NextRow = NextRow + LastRow - FirstRow + 1
            '�R�s�[����Ԃ�����
            Application.CutCopyMode = False
            '�\�t�����̃u�b�N��ۑ������ɕ���
            .Close False
        End With
    Next a
    
    With Application
        .ScreenUpdating = True  '��ʍX�V�L��
        .EnableEvents = True    '�C�x���g�L����
        .Calculation = xlCalculationAutomatic   '�v�Z������
    End With
    
    MsgBox "�W�񊮗�"
    
End Sub


