Attribute VB_Name = "Module1"
Option Explicit
 
 Sub SplitFiles()

    Dim FN1 As String          'このブックのファイル名
    Dim MacroWS As Worksheet   'このブックのシート
    Dim Wb_new As Workbook     '分割後のブック
    Dim WS As Worksheet        '分割元データのシート
    Dim rowsData As Long       '分割元ブックのデータ数（行数）
    Dim colsData As Long       '分割元ブックのデータ数（列数）
    Dim R_Data2 As Long        '分割元ブックのデータ開始行（実データ開始行）
    Dim Ko As Long             '分割ファイル数（係の数）

    Dim Wb_Data As Workbook    '1. 分割元ブック
    Dim R_Data1 As Long        '2. 分割元ブックのデータ開始行（タイトル行）
    Dim Path As String         '3. 分割データ保存先
    Dim C_Group As String      '4. グループ対象列
    Dim My_Group As String     '5. 自係名
    Dim Uni_Word As String     '6. ユニークワード
    Dim FN2 As String          '7. 分割後ブックのファイル名
    Dim PSW As String          '8. 読み取りパスワード
    
    '値をセット
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
    
    Application.ScreenUpdating = False  '画面を固定して高速化
    
    '最終行、最終列の取得
    Wb_Data.Activate
    rowsData = WS.Cells(Rows.Count, 1).End(xlUp).Row
    colsData = WS.Cells(R_Data1, Columns.Count).End(xlToLeft).Column
    
    '係名でソート
    WS.Range(Rows(R_Data1), Rows(rowsData)).Sort _
        Key1:=Range(C_Group & R_Data1), _
        Order1:=xlAscending, _
        Header:=xlYes, _
        Orientation:=xlTopToBottom
    
    '係名ごとにファイルを分割し保存
    Do
        '元ファイルのデータ開始行（項目行）をコピーし、新規エクセルブックに貼り付け
        Wb_Data.Activate
        WS.Range(Cells(R_Data1, 1), Cells(R_Data1, colsData)).Copy
        Workbooks.Add
        ActiveSheet.Paste Range("A2") '2行目以降にデータを記載（1行目はユニークワード記載用に空けておく）
        Set Wb_new = ActiveWorkbook
        
        '１係分のみ抽出し、ファイル名を設定して保存
        Wb_Data.Activate
        Ko = WorksheetFunction.CountIf(Columns(C_Group), Cells(R_Data2, C_Group)) '１係分のデータ数を算出
        Range(Cells(R_Data2, "A"), Cells(R_Data2 + Ko - 1, colsData)).Copy        '１係分のデータ数分コピー
        Wb_new.Activate
        ActiveSheet.Paste Range("A3")                                             '新規ブックの3行目以下に値貼り付け
        If Cells(3, C_Group) = My_Group Then
            Range("A1").Value = Uni_Word                                          '自係のファイルのみ、A1セルにユニークワードを記載
        End If
        Wb_new.SaveAs FileName:=Path & Cells(3, C_Group) & FN2 & ".xlsx", _
        Password:=PSW                                                             '指定したフォルダーに保存
        
        Wb_new.Close

        R_Data2 = R_Data2 + Ko
        Loop While Cells(R_Data2, C_Group) <> ""
    
    MsgBox "分割処理完了"

    Application.ScreenUpdating = True

End Sub

