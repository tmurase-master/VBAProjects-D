Attribute VB_Name = "Module1"
Sub ���̃t�@�C������W��()

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
