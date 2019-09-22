Attribute VB_Name = "Module1"
Option Explicit

Sub CSV読込マクロ()
    
    Dim File_Path As String
    Dim CSV_Array As Variant
    
    '-----------------------------------------------
    '☆ チェック作業（本格的に作業する前に確認）
    If Not FnGetTitleList(CSV_Array) Then Exit Sub  'CSVの設定情報の読み取り。エラーがあればマクロを終了する
    If Not FnFilePicker(File_Path) Then Exit Sub    'ファイルの読み込み。パスを取得できない場合はマクロを終了する
    
    
    '-----------------------------------------------
    '☆ 初期作業
    MacroMode = True        '負荷軽減
    
    Call CSV_Initialaize
    
    '-----------------------------------------------
    '☆ データの読み取り
    
    Call CSV_read_Macro(File_Path, CSV_Array)
    
    
    '-----------------------------------------------
    '☆ 終了処理
    
    Worksheets("CSV").Activate
    MacroMode = False        '負荷軽減　戻し
    
    MsgBox "作業が終了しました。" & vbCrLf & _
            "「CSV」シートを確認してください。"
    
End Sub


'====================================================================
'　CSVデータ各列の設定値を取得する関数
'　返り値：True/False (Boolean型)
'　　　　　True ：設定値を取得できた場合
'　　　　　False：設定値を取得できなかった場合（初期値）
'　引数　：CSV_Array　設定値の配列（long型の1次元配列）
'　　　　　　　　　　(ByRef指定により)取得した設定値を格納して返す
'====================================================================

Private Function FnGetTitleList(ByRef CSV_Array As Variant) As Boolean
    
    On Error GoTo Err_Data      '途中で不明なエラーが発生した場合、エラーメッセージを表示
    
    '-----------------------------------------------
    '☆ データの読み取り
'    Dim sh As Worksheet
'    Set sh = Worksheets("読込設定")
    
    Dim 設定リスト As Variant
    設定リスト = Worksheets("読込設定").Range("B6").CurrentRegion
    
    Dim 件数 As Long
    件数 = UBound(設定リスト, 1) - 1        'タイトル行を除いた件数
    
    If 件数 < 1 Then            '件数が0件の場合エラー処理
        MsgBox "[読み込み設定]が設定されていません" & vbCrLf & _
                "「読込設定」シートの設定をしてください", _
                vbCritical, _
                "「読込設定」シート 件数0　エラー"
        Exit Function
    End If
    
    '-----------------------------------------------
    '☆ データの書き出し処理
    Dim i As Long
    Dim tmp() As Long
    ReDim tmp(件数 - 1)
    
    For i = 1 To 件数
        tmp(i - 1) = CLng(Left(設定リスト(i + 1, 2), 1))
    Next i
    
    CSV_Array = tmp         '配列の格納
    FnGetTitleList = True   '成功結果を返す
    
    
    Exit Function           '処理の終了
    
Err_Data:
    MsgBox "[読み込み設定]を読み取り時に不明なエラーが発生しました" & vbCrLf & _
            "「読込設定」シートの内容を確認してください", _
            vbCritical, _
            "「読込設定」シート　読み取り　エラー"
    
End Function


'====================================================================
'　ファイルのパスを取得する関数
'　返り値：True/False (Boolean型)
'　　　　　True ：パスを取得できた場合
'　　　　　False：パスを取得できなかった場合
'　引数　：File_Path　ファイルのパス
'　　　　　　　　　　(ByRef指定により)取得したパスを格納して返す
'　固有変数：fFlg　実行済みフラグ
'====================================================================

Private Function FnFilePicker(ByRef File_Path As String) As Boolean
    
    Static fFlg As Boolean
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        '初回のみ、このファイルのパスを初期フォルダとする
        If Not fFlg Then
            .InitialFileName = ThisWorkbook.Path
            fFlg = True
        End If
        
        With .Filters   '選択可能ファイルの絞り込み
            .Clear
            .Add "CSVファイル", "*.csv,*.txt", 1
            .Add "全てのファイル", "*.*", 2
        End With
        
        'ダイアログよりパスを取得。成否結果を変数の結果として返す
        If .Show = True Then
            File_Path = .SelectedItems(1)
            FnFilePicker = True
        End If
    End With
    
End Function

'====================================================================
'　CSVシートの初期化
'　（必要に応じて内容追加）
'====================================================================

Private Sub CSV_Initialaize()
    Worksheets("CSV").Cells.Delete         'シートの中身を全て削除
End Sub



'====================================================================
'　CSVを読み込む処理をするプロシージャ（使いまわしが多いため独立)
'　引数　：File_Path　CSVファイルのパス
'          CSV_Array　CSV読込の配列
'　変数　：fFlg　実行済みフラグ
'====================================================================

Private Sub CSV_read_Macro(File_Path As String, CSV_Array)
    Dim sh As Worksheet
    Set sh = Worksheets("CSV")
    
    With sh.QueryTables.Add( _
        Connection:="TEXT;" & File_Path, _
        Destination:=sh.Range("A1"))            'Connection:読み込みファイル、Destination:貼付け先
        
        .Name = "temp"                          '今回の読み込み操作の名称（最後に削除するのでなんでもよい）
        .AdjustColumnWidth = True               '列幅の自動設定
        .TextFilePlatform = 932                 '文字コード：932 SJIS
        .TextFileCommaDelimiter = True          'カンマ区切り
        .TextFileColumnDataTypes = CSV_Array    '1=数字、2=文字列
        .Refresh BackgroundQuery:=False         'バックグラウンド処理(False:しない。バックグラウンド処理すると読み込み前にマクロが次へ進むため)
        
        .Delete                                 'クエリの削除(データ接続を消す)
    End With
    
    Set sh = Nothing
End Sub


'====================================================================
'　マクロ実行時によく使う設定をまとめて処理
'　引数　：Flag　True/False (Boolean型)
'====================================================================

Property Let MacroMode(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag            'イベントの実行・停止
        .ScreenUpdating = Not Flag          '画面更新の実行・停止
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)   '再計算の実行・停止
    End With
End Property

