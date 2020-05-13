Attribute VB_Name = "Module1"
Sub 他のファイルから集約()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
 
    Dim Search_sname As String  '検索結果が記録されているシート名
    Search_sname = "result"
    
    Dim Sheetname As String   'PJ資料内のシート名を格納
    Sheetname = "Sheet1"

    Dim Actual_r As String    '集約（コピペ）範囲を格納
    Actual_r = "A2: H2"

    Dim resultnum As Long   '検索結果数
    Dim a As Long   'ループ用変数
    
    Dim Filepath() As String      'PJ資料名を格納


    '配列の初期化
    Erase Filepath

    '検索結果のファイルパスをFilePath配列へ格納
    With ThisWorkbook.Worksheets(Search_sname)
        resultnum = 0
    
        '配列への格納
        Do Until .Cells(resultnum + 3, 2) = ""
            ReDim Preserve Filepath(resultnum + 1)
            Filepath(resultnum) = .Cells(resultnum + 3, 2)
            resultnum = resultnum + 1
        Loop
    End With
            

    '値の集約
    a = 0

    Do Until Filepath(a) = ""
        With Workbooks.Open(Filepath(a))
            '「貼付元」シートの対象範囲をコピー
            .Worksheets(Sheetname).Range(Actual_r).Copy
            'このブックの選択しているセルへ値貼り付け
            ThisWorkbook.Worksheets("集約").Cells(a + 2, 1).PasteSpecial xlPasteAll
            'コピー中状態を解除
            Application.CutCopyMode = False
            'ブックを保存せずに閉じる
            .Close False
            a = a + 1
    
        End With
    Loop
    
    MsgBox "集約完了"
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
