Attribute VB_Name = "Module1"
Sub ���̃t�@�C������W��()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
 
    Dim Search_sname As String  '�������ʂ��L�^����Ă���V�[�g��
    Search_sname = "result"
    
    Dim Sheetname As String   'PJ�������̃V�[�g�����i�[
    Sheetname = "Sheet1"

    Dim Actual_r As String    '�W��i�R�s�y�j�͈͂��i�[
    Actual_r = "A2: H2"

    Dim resultnum As Long   '�������ʐ�
    Dim a As Long   '���[�v�p�ϐ�
    
    Dim Filepath() As String      'PJ���������i�[


    '�z��̏�����
    Erase Filepath

    '�������ʂ̃t�@�C���p�X��FilePath�z��֊i�[
    With ThisWorkbook.Worksheets(Search_sname)
        resultnum = 0
    
        '�z��ւ̊i�[
        Do Until .Cells(resultnum + 3, 2) = ""
            ReDim Preserve Filepath(resultnum + 1)
            Filepath(resultnum) = .Cells(resultnum + 3, 2)
            resultnum = resultnum + 1
        Loop
    End With
            

    '�l�̏W��
    a = 0

    Do Until Filepath(a) = ""
        With Workbooks.Open(Filepath(a))
            '�u�\�t���v�V�[�g�̑Ώ۔͈͂��R�s�[
            .Worksheets(Sheetname).Range(Actual_r).Copy
            '���̃u�b�N�̑I�����Ă���Z���֒l�\��t��
            ThisWorkbook.Worksheets("�W��").Cells(a + 2, 1).PasteSpecial xlPasteAll
            '�R�s�[����Ԃ�����
            Application.CutCopyMode = False
            '�u�b�N��ۑ������ɕ���
            .Close False
            a = a + 1
    
        End With
    Loop
    
    MsgBox "�W�񊮗�"
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
