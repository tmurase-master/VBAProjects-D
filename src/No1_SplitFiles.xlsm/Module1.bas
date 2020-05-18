Attribute VB_Name = "Module1"
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
    Set MacroWS = Workbooks(FN1).Worksheets(1)
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
        Wb_new.SaveAs FileName:=Path & Cells(3, C_Group) & FN2 & ".xlsx", _
        Password:=PSW                                                             '�w�肵���t�H���_�[�ɕۑ�
        
        Wb_new.Close

        R_Data2 = R_Data2 + Ko
        Loop While Cells(R_Data2, C_Group) <> ""
    
    MsgBox "������������"

    Application.ScreenUpdating = True

End Sub

