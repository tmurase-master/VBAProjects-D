Attribute VB_Name = "Module1"
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

Dim LogCount As Long

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
 
 '��ʍX�V���I�t�ɂ���
 Application.ScreenUpdating = False
 
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
 
 '��ʍX�V���I���ɂ���
 Application.ScreenUpdating = True
 
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

Public Sub FileAggregation()

    With Application
        .ScreenUpdating = False '��ʍX�V����
        .EnableEvents = False   '�C�x���g�}�~
        .Calculation = xlCalculationManual  '�v�Z�蓮��
    End With
 
    Dim Search_sname As String  '�������ʂ��L�^����Ă���V�[�g��
    Search_sname = "Result"

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
            .Worksheets(Filepath(a, 1)).Range(Cells(FirstRow, 1), Cells(LastRow, LastCol)).Copy
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
    
End Sub
