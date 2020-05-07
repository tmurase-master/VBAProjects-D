Attribute VB_Name = "Module1"
'---------------------------------
'�쐬���F2020/04/30
'�쐬�ҁF���c
'���O�������
'�E�����L�[���[�h�͑Ώۃt�@�C����A1�s�Ɋi�[����Ă�����̂Ƃ���B
'�E�����Ώۃt�@�C���̓G�N�Z���t�@�C���i.xls/.xls*)�݂̂Ƃ���B
'�E�����Ώۂ͎w�肵���t�H���_�������i�T�u�t�H���_�͊܂܂Ȃ��j�̑S�G�N�Z���t�@�C���Ƃ���B
'���C���v�b�g�E�A�E�g�v�b�g��
'�C���v�b�g�@�F�����i�����Ώۃt�H���_�����͖{�t�@���N�V�������Ŏ擾����j
'�A�E�g�v�b�g�F�w��L�[���[�h���܂ރt�@�C�����i��΃p�X�Ŗ{���[�N�u�b�N����"Result"�V�[�g�֏������ށj
'---------------------------------

Sub SearchFiles()
 Dim ResultTab_Row As Long 'Result�V�[�g�̏����Ώۍs�p�ϐ��i�f�t�H���g�l�̓t�H�[�}�b�g�ɍ��킹�āu2�v�Ŏw�肷��j
 Dim LogTab_Row As Long 'Log�V�[�g�̏����Ώۍs�p�ϐ��i�f�t�H���g�l�̓t�H�[�}�b�g�ɍ��킹�āu2�v�Ŏw�肷��j
 Dim SearchWord As String '�����L�[���[�h�p�ϐ��i�C���v�b�g�{�b�N�X�Ŏ擾����j
 Dim SearchFolder As String '�����Ώۂ̃t�H���_���p�ϐ��i�_�C�����O�Ŏ擾�j
 Dim Get_Filename As String '�擾�����t�@�C�������ꎞ�I�Ɋi�[����ϐ�
 Dim LoopCount_1 'Loop�J�E���g�p�ϐ�
 Dim LoopCount_2 'Loop�J�E���g�p�ϐ�
 Dim Check_sheet As Worksheet '���[�N�V�[�g�����p�ϐ�
 Dim SearchFilename As String '�����ΏۂƂ��ăt�@�C���I�[�v������ۂɃt�@�C�������i�[����ϐ�
 
 ResultTab_Row = 3 '�����l���Q�Őݒ�
 LogTab_Row = 1  '�����l���Q�Őݒ�
 
 '�����Ώۂ̃t�H���_���w��
 '��U�_�~�[�Œ��ڃp�X�w��i���ƂŃ_�C�����O�w��Ƃ��đΉ��ł���悤����������j
 SearchFolder = "C:\Users\SX2\Desktop\macro_dev\budget\"
 
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
 ThisWorkbook.Worksheets("Log").Range("B3:B102").Clear
 ThisWorkbook.Worksheets("Log").Columns("A").Clear

 '�w�肵���t�H���_�����ɂ���G�N�Z���t�@�C�������擾���ALog�V�[�g�֏����o��
 Get_Filename = Dir(SearchFolder & "*.xls") '�w��t�H���_�z����.xls���܂ރt�@�C�����擾�i1�t�@�C�����j
 LoopCount_1 = LogTab_Row '���[�v�p�ϐ��Ƀ��O�V�[�g�����l���Z�b�g

 Do While Get_Filename <> "" '�w�肵���t�H���_������.xls���܂ރt�@�C������S�Ď擾����܂Ń��[�v
    ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) = SearchFolder & Get_Filename '�擾�����t�@�C������Log�V�[�g�֏�������
    LoopCount_1 = LoopCount_1 + 1 '���s�֐i��
    Get_Filename = Dir() '�c��̃t�@�C�������擾�i1�t�@�C�����j
 Loop
 
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
