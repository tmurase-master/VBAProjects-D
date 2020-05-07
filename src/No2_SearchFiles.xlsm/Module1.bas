Attribute VB_Name = "Module1"
'---------------------------------
'作成日：2020/04/30
'作成者：福田
'＜前提条件＞
'・検索キーワードは対象ファイルのA1行に格納されているものとする。
'・検索対象ファイルはエクセルファイル（.xls/.xls*)のみとする。
'・検索対象は指定したフォルダ名直下（サブフォルダは含まない）の全エクセルファイルとする。
'＜インプット・アウトプット＞
'インプット　：無し（検索対象フォルダ名等は本ファンクション内で取得する）
'アウトプット：指定キーワードを含むファイル名（絶対パスで本ワークブック内の"Result"シートへ書き込む）
'---------------------------------

Sub SearchFiles()
 Dim ResultTab_Row As Long 'Resultシートの処理対象行用変数（デフォルト値はフォーマットに合わせて「2」で指定する）
 Dim LogTab_Row As Long 'Logシートの処理対象行用変数（デフォルト値はフォーマットに合わせて「2」で指定する）
 Dim SearchWord As String '検索キーワード用変数（インプットボックスで取得する）
 Dim SearchFolder As String '検索対象のフォルダ名用変数（ダイヤログで取得）
 Dim Get_Filename As String '取得したファイル名を一時的に格納する変数
 Dim LoopCount_1 'Loopカウント用変数
 Dim LoopCount_2 'Loopカウント用変数
 Dim Check_sheet As Worksheet 'ワークシート処理用変数
 Dim SearchFilename As String '検索対象としてファイルオープンする際にファイル名を格納する変数
 
 ResultTab_Row = 3 '初期値を２で設定
 LogTab_Row = 1  '初期値を２で設定
 
 '検索対象のフォルダを指定
 '一旦ダミーで直接パス指定（あとでダイヤログ指定として対応できるよう処理を入れる）
 SearchFolder = "C:\Users\SX2\Desktop\macro_dev\budget\"
 
 '検索キーワードをインプットボックスで取得
 Do While SearchWord = "" 'キーワードを1文字以上入力されるまでループ
    SearchWord = InputBox("検索するキーワードを入力してください", "キーワード入力", "") 'インプットボックスで取得
    If SearchWord = "" Then  '何もキーワードを入力されなかった場合
        MsgBox "キーワードを1文字以上入力してください。" 'エラーメッセージを表示する
    End If
 Loop
 
 '画面更新をオフにする
 Application.ScreenUpdating = False
 
 '取得したキーワードをResultシートに書き込む
 ThisWorkbook.Worksheets("Result").Range("B1") = SearchWord

 'Result/Logシートをクリアする
 ThisWorkbook.Worksheets("Log").Range("B3:B102").Clear
 ThisWorkbook.Worksheets("Log").Columns("A").Clear

 '指定したフォルダ直下にあるエクセルファイル名を取得し、Logシートへ書き出す
 Get_Filename = Dir(SearchFolder & "*.xls") '指定フォルダ配下の.xlsを含むファイルを取得（1ファイル分）
 LoopCount_1 = LogTab_Row 'ループ用変数にログシート初期値をセット

 Do While Get_Filename <> "" '指定したフォルダ直下の.xlsを含むファイル名を全て取得するまでループ
    ThisWorkbook.Worksheets("Log").Range("A" & LoopCount_1) = SearchFolder & Get_Filename '取得したファイル名はLogシートへ書き込む
    LoopCount_1 = LoopCount_1 + 1 '次行へ進む
    Get_Filename = Dir() '残りのファイル名を取得（1ファイル分）
 Loop
 
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
        LoopCount_2 = LoopCount_2 + 1 'Resultシートの入力行を進める
        Exit For 'いずれか１つのシートでキーワードが合致した場合は当該ファイルへの検索処理を終了する。
    End If
    Next Check_sheet '次のシートへ移動する
    
    ActiveWorkbook.Close '処理が終わったファイルを閉じる
    LoopCount_1 = LoopCount_1 + 1 'Logシートの行を一つ進める
 Loop
 
 '画面更新をオンにする
 Application.ScreenUpdating = True
 
 '完了メッセージを表示する
 MsgBox "処理が完了しました"
 
End Sub
