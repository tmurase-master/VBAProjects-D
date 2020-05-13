Attribute VB_Name = "Module1"
Option Explicit

 Sub SplitFiles()

    Dim MacroB As Worksheet    'このブックのシート
    Dim Wb_new As Workbook     '分割データ保存ブック
    Dim Wb_Data As Workbook    '1. 分割元ブック
    Dim Ws As String           '2. 分割元シート名
    Dim R_Data As Integer      '3. 分割元ブックのデータ開始行
    Dim Path As String         '4. 分割データ保存先
    Dim C_Group As String      '5. グループ対象列
    Dim C_Copy As String       '6. コピーデータ右端列
    Dim R_Copy As Integer      '7. コピーデータ最終行
    Dim FN As String           '8. 保存ブック日付の表示形式
    Dim PSW As String          '9. 読み取りパスワード

    Dim Ko As Integer    'グループの件数

    Set MacroB = Workbooks("No1_SplitFiles.xlsm").Worksheets(1)   'このブックのシート
    Set Wb_Data = Workbooks(MacroB.Range("C3").Value)    '分割元のブック名
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
        Worksheets(Ws).Range(Cells(1, 1), Cells(1, C_Copy)).Copy    '1行目の項目名コピー
        Workbooks.Add
        ActiveSheet.Paste Range("A1")    '新規ブックに貼り付け
        Set Wb_new = ActiveWorkbook

        Wb_Data.Activate
        Ko = WorksheetFunction.CountIf(Columns(C_Group), Cells(R_Data, C_Group)) 'グループの件数を算出
        Range(Cells(R_Data, "A"), Cells(R_Data + Ko - 1, C_Copy)).Copy    'グループ件数分コピー
        Wb_new.Activate
        ActiveSheet.Paste Range("A2")    '新規ブック項目の下に貼り付け
        Wb_new.SaveAs FileName:=Path & Cells(2, C_Group) & FN & ".xlsx", _
        Password:=PSW    '指定したフォルダーに保存
        Wb_new.Close

        R_Data = R_Data + Ko

        Loop While Cells(R_Data, C_Group) <> ""
    MsgBox "完了！"

    Application.ScreenUpdating = True

End Sub

