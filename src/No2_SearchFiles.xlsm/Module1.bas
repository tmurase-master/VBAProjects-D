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
 
 ResultTab_Row = 3 '�����l���Q�Őݒ�
 LogTab_Row = 1  '�����l���Q�Őݒ�
 
 '�����Ώۂ̃t�H���_���w��
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
        MsgBox "�L�[���[�h��1�����ȏ���͂��Ă��������B" '�G���[���b�Z�[�W��\������
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
 Dim LoopCount_1 As Long
 Dim CheckName As String
 
 LoopCount_1 = 1
 
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Folder = FSO.GetFolder(TargetDir)
 
 For Each SubFolder In Folder.SubFolders
    FolderSearch SubFolder.Path
 Next SubFolder
 
 For Each Filename In Folder.Files
    TmpName = Filename
    CheckName = Mid(TmpName, InStrRev(TmpName, "\") + 1)
    
    If CheckName <> ThisWorkbook.Name Then '�{�}�N���t�@�C���i���ꖼ�̃t�@�C���j�͌����ΏۊO�Ƃ���B
        If TmpName Like "*.xls*" Then
            ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) = TmpName '�擾�����t�@�C������Log�V�[�g�֏�������
            LoopCount_1 = LoopCount_1 + 1 '���s�֐i��
        End If
    End If
 Next Filename

End Sub
