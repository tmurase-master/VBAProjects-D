Attribute VB_Name = "Module1"
Sub 他のファイルから集約()

    With Application
        .ScreenUpdating = False '画面更新無効
        .EnableEvents = False   'イベント抑止
        .Calculation = xlCalculationManual  '計算手動化
    End With
 
    Dim Search_sname As String  '検索結果が記録されているシート名
    Search_sname = "result"

    Dim resultnum As Long   '検索結果数
    Dim a As Long   'ループ用変数
    Dim FirstRow As Long  '貼付け元シートのコピー範囲初期行数を格納
    Dim NextRow As Long   '集約シートへの貼付け先の行数を格納
    Dim Filepath() As String      '（PJ資料名, シート名）を格納
    Dim LastRow As Long '貼付け元の予算シート最終行を格納
    Dim LastCol As Long '貼付け元の予算シート最終列を格納


    '検索結果のファイルパスをFilePath配列へ格納
    With ThisWorkbook.Worksheets(Search_sname)
        resultnum = 0
    
        '検索結果数の確認
        Do Until .Cells(resultnum + 3, 2) = ""
            resultnum = resultnum + 1
        Loop
        
        '検索結果が0件の場合処理終了
        If resultnum = 0 Then
            MsgBox "検索結果が0件のため処理を終了します"
            End
        End If
         
        resultnum = resultnum - 1
        ReDim Filepath(resultnum, 1) '配列の要素数を設定
        
        For a = 0 To resultnum
            Filepath(a, 0) = .Cells(a + 3, 2) 'ファイルパスの格納
            Filepath(a, 1) = .Cells(a + 3, 3) 'シート名の格納
        Next a
    End With
            
    FirstRow = 3  '貼付け元シートのコピー範囲初期行数を設定
    NextRow = 2 '集約シートの貼付け先初期行数を設定
    
    '集約シートへコピー
    For a = 0 To resultnum
        With Workbooks.Open(Filepath(a, 0))
            '貼り付け元の範囲を取得
            LastRow = .Worksheets(Filepath(a, 1)).Cells(Rows.Count, 1).End(xlUp).Row
            LastCol = .Worksheets(Filepath(a, 1)).Cells(LastRow, Columns.Count).End(xlToLeft).Column
            '貼り付け元のコピー
            .Worksheets(Filepath(a, 1)).Range(Worksheets(Filepath(a, 1)).Cells(FirstRow, 1), _
                                              Worksheets(Filepath(a, 1)).Cells(LastRow, LastCol)).Copy
            'このブックの集約セルへ値貼り付け
            ThisWorkbook.Worksheets("集約").Cells(NextRow, 1).PasteSpecial xlPasteAll
            '次回貼付け位置の設定
            NextRow = NextRow + LastRow - FirstRow + 1
            'コピー中状態を解除
            Application.CutCopyMode = False
            '貼付け元のブックを保存せずに閉じる
            .Close False
        End With
    Next a
    
    With Application
        .ScreenUpdating = True  '画面更新有効
        .EnableEvents = True    'イベント有効化
        .Calculation = xlCalculationAutomatic   '計算自動化
    End With
    
    MsgBox "集約完了"
    
End Sub
