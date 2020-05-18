Attribute VB_Name = "Module1"
'---------------------------------
'作成日：2020/04/30
'作成者：福田
'＜使用環境＞
'タブメニューの「ツール」－「参照設定」で「Microsoft Scripting Runtime」を導入（チェック）すること。
'＜前提条件＞
'・検索キーワードは対象ファイルのA1行に格納されているものとする。
'・検索対象ファイルはエクセルファイル（.xls/.xls*)のみとする。
'・検索対象は指定したフォルダ名直下（サブフォルダは含まない）の全エクセルファイルとする。
'＜インプット・アウトプット＞
'インプット　：無し（検索対象フォルダ名等は本ファンクション内で取得する）
'アウトプット：指定キーワードを含むファイル名（絶対パスで本ワークブック内の"Result"シートへ書き込む）
'＜アップデート＞
'5/18アップデート：特定環境対応、サブフォルダ対応他
'---------------------------------

Dim LogCount As Long

Sub SearchFiles()
 Dim ResultTab_Row As Long 'Resultシートの処理対象行用変数（デフォルト値はフォーマットに合わせて「2」で指定する）
 Dim LogTab_Row As Long 'Logシートの処理対象行用変数（デフォルト値はフォーマットに合わせて「2」で指定する）
 Dim SearchWord As String '検索キーワード用変数（インプットボックスで取得する）
 Dim SearchFolder As String '検索対象のフォルダ名用変数（ダイヤログで取得）
 Dim Get_Foldername As String 'ファイルダイヤログにて取得した対象フォルダパス
 Dim Get_Foldername2 As String 'フォルダパス加工用変数
 Dim StrLen As Long '文字列長計算用変数
 Dim Get_Filename As String '取得したファイル名を一時的に格納する変数
 Dim LoopCount_1 'Loopカウント用変数
 Dim LoopCount_2 'Loopカウント用変数
 Dim Check_sheet As Worksheet 'ワークシート処理用変数
 Dim SearchFilename As String '検索対象としてファイルオープンする際にファイル名を格納する変数
 
 ResultTab_Row = 3 '初期値を３で設定
 LogTab_Row = 1  '初期値を１で設定
 LogCount = 1
 
 
 '検索対象のフォルダを指定
 MsgBox "検索対象フォルダを選択してください。選択したフォルダ配下にあるサブフォルダも検索対象となります。"
 With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
        Get_Foldername = .SelectedItems(1)
    Else
        MsgBox "終了します。再実行してください。"
        Exit Sub
    End If
 End With

 '特定環境での動作保証
 If Left(Get_Foldername, 11) = "http://prdo" Then
    StrLen = Len(Get_Foldername)
    Get_Foldername2 = Right(Get_Foldername, StrLen - 28)
    SearchFolder = "G:" & Get_Foldername2
 Else
    SearchFolder = Get_Foldername
 End If
 
 SearchFolder = SearchFolder & "\"
 
 '検索キーワードをインプットボックスで取得
 Do While SearchWord = "" 'キーワードを1文字以上入力されるまでループ
    SearchWord = InputBox("検索するキーワードを入力してください", "キーワード入力", "") 'インプットボックスで取得
    If SearchWord = "" Then  '何もキーワードを入力されなかった場合
        MsgBox "キーワードを1文字以上入力してください。なお、キャンセルはできません。" 'エラーメッセージを表示する
    End If
 Loop
 
 '画面更新をオフにする
 Application.ScreenUpdating = False
 
 '取得したキーワードをResultシートに書き込む
 ThisWorkbook.Worksheets("Result").Range("B1") = SearchWord

 'Result/Logシートをクリアする
 ThisWorkbook.Worksheets("Result").Range("B3:C102").ClearContents
 ThisWorkbook.Worksheets("Log").Columns("A").ClearContents

'5/18 サブフォルダ検索に伴い廃止（SearchFolder関数にて対応）
' 指定したフォルダ直下にあるエクセルファイル名を取得し､Logシートへ書き出す
' Get_Filename = Dir(SearchFolder & "*.xls") '指定フォルダ配下の.xlsを含むファイルを取得（1ファイル分）
' LoopCount_1 = LogTab_Row 'ループ用変数にログシート初期値をセット
'
' Do While Get_Filename <> "" '指定したフォルダ直下の.xlsを含むファイル名を全て取得するまでループ
'    If Get_Filename <> ThisWorkbook.Name Then '本マクロファイル（同一名称ファイル）は検索対象外とする。
'        ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) = SearchFolder & Get_Filename '取得したファイル名はLogシートへ書き込む
'        LoopCount_1 = LoopCount_1 + 1 '次行へ進む
'    End If
'    Get_Filename = Dir() '残りのファイル名を取得（1ファイル分）
' Loop
 
 Call FolderSearch(SearchFolder)
 
 'Logシートに書き込まれたファイルを全て開いてチェックする
 LoopCount_1 = LogTab_Row
 LoopCount_2 = ResultTab_Row 'Resultシートの初期値をセット
 Do While ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) <> ""  'Logシートに記載されたファイルへの処理が全て終わるまでループ
    Workbooks.Open ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) 'エクセルファイルを開く（Logシートの上から順）
    
    '各シートのA1セルにキーワードが含まれているかチェックし、含まれていればファイル名をResultシートに書き込む
    For Each Check_sheet In ActiveWorkbook.Worksheets '全シートへの処理が完了するまでループ
        Check_sheet.Activate 'Sheetをアクティベーションする（全シート処理のために必要な処理）
        If ActiveSheet.Range("A1") = SearchWord Then 'A1セルに記載されたキーワードと入力した検索ワードが一致した場合
            ThisWorkbook.Worksheets("Result").Range("B" & LoopCount_2) = ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) 'Resultシートにファイル名（絶対パス）を入力
            ThisWorkbook.Worksheets("Result").Range("C" & LoopCount_2) = ActiveSheet.Name 'Resultシートにファイル名（絶対パス）を入力
            LoopCount_2 = LoopCount_2 + 1 'Resultシートの入力行を進める
            Exit For 'いずれか１つのシートでキーワードが合致した場合は当該ファイルへの検索処理を終了する。
        End If
    Next Check_sheet '次のシートへ移動する
    
    ActiveWorkbook.Close '処理が終わったファイルを閉じる
    LoopCount_1 = LoopCount_1 + 1 'Logシートの行を一つ進める
 Loop
 
 '貼り付け対象ファイルが存在しなかった場合は、貼り付け処理前に終了させる。
 If LoopCount_2 = 3 Then
    Application.ScreenUpdating = True
    MsgBox "指定されたキーワードに合致するファイルは存在しませんでした。処理を終了します。"
    Exit Sub
 End If
 
 Call FileAggregation 'ファイル書き込み処理を実行
 
 '画面更新をオンにする
 Application.ScreenUpdating = True
 
 '完了メッセージを表示する
 MsgBox "処理が完了しました"
 
End Sub
Public Sub FolderSearch(TargetDir As String) 'サブフォルダーを含む対象ファイル（エクセルファイル）一覧取得
 Dim FSO As Object
 Dim Folder As Object
 Dim SubFolder As Object
 Dim Filename As Object
 Dim TmpName As String
 Dim CheckName As String
 
 
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Folder = FSO.GetFolder(TargetDir)
 
 For Each SubFolder In Folder.SubFolders
    Call FolderSearch(SubFolder.Path)
 Next SubFolder
 
 For Each Filename In Folder.Files
    TmpName = Filename
    CheckName = Mid(TmpName, InStrRev(TmpName, "\") + 1)
    
    If CheckName <> ThisWorkbook.Name Then '本マクロファイル（同一名称ファイル）は検索対象外とする。
        If TmpName Like "*.xls*" Then
            ThisWorkbook.Worksheets("Log").Range("A" & LogCount) = TmpName '取得したファイル名はLogシートへ書き込む
            LogCount = LogCount + 1 '次行へ進む
        End If
    End If
 Next Filename

End Sub

Public Sub FileAggregation()

    With Application
        .ScreenUpdating = False '画面更新無効
        .EnableEvents = False   'イベント抑止
        .Calculation = xlCalculationManual  '計算手動化
    End With
 
    Dim Search_sname As String  '検索結果が記録されているシート名
    Search_sname = "Result"

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
            .Worksheets(Filepath(a, 1)).Range(Cells(FirstRow, 1), Cells(LastRow, LastCol)).Copy
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
    
End Sub
