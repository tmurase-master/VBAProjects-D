Attribute VB_Name = "Module1"
Option Explicit

 Sub SplitFiles()

    Dim MacroB As Worksheet    '���̃u�b�N�̃V�[�g
    Dim Wb_new As Workbook     '�����f�[�^�ۑ��u�b�N
    Dim Wb_Data As Workbook    '1. �������u�b�N
    Dim Ws As String           '2. �������V�[�g��
    Dim R_Data As Integer      '3. �������u�b�N�̃f�[�^�J�n�s
    Dim Path As String         '4. �����f�[�^�ۑ���
    Dim C_Group As String      '5. �O���[�v�Ώۗ�
    Dim C_Copy As String       '6. �R�s�[�f�[�^�E�[��
    Dim R_Copy As Integer      '7. �R�s�[�f�[�^�ŏI�s
    Dim FN As String           '8. �ۑ��u�b�N���t�̕\���`��
    Dim PSW As String          '9. �ǂݎ��p�X���[�h

    Dim Ko As Integer    '�O���[�v�̌���

    Set MacroB = Workbooks("No1_SplitFiles.xlsm").Worksheets(1)   '���̃u�b�N�̃V�[�g
    Set Wb_Data = Workbooks(MacroB.Range("C3").Value)    '�������̃u�b�N��
    Ws = MacroB.Range("C4")
    R_Data = MacroB.Range("C5")
    Path = MacroB.Range("C6")
    C_Group = MacroB.Range("C7")
    C_Copy = MacroB.Range("C8")
    R_Copy = MacroB.Range("C9")
    FN = MacroB.Range("C10")
    PSW = MacroB.Range("C11")

    Application.ScreenUpdating = False
    
    Do
        Wb_Data.Activate
        Worksheets(Ws).Range(Cells(1, 1), Cells(1, C_Copy)).Copy    '1�s�ڂ̍��ږ��R�s�[
        Workbooks.Add
        ActiveSheet.Paste Range("A1")    '�V�K�u�b�N�ɓ\��t��
        Set Wb_new = ActiveWorkbook

        Wb_Data.Activate
        Ko = WorksheetFunction.CountIf(Columns(C_Group), Cells(R_Data, C_Group)) '�O���[�v�̌������Z�o
        Range(Cells(R_Data, "A"), Cells(R_Data + Ko - 1, C_Copy)).Copy    '�O���[�v�������R�s�[
        Wb_new.Activate
        ActiveSheet.Paste Range("A2")    '�V�K�u�b�N���ڂ̉��ɓ\��t��
        Wb_new.SaveAs FileName:=Path & Cells(2, C_Group) & FN & ".xlsx", _
        Password:=PSW    '�w�肵���t�H���_�[�ɕۑ�
        Wb_new.Close

        R_Data = R_Data + Ko

        Loop While Cells(R_Data, C_Group) <> ""
    MsgBox "�����I"

    Application.ScreenUpdating = True

End Sub

