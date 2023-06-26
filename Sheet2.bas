VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub 成績情報登録ボタンクリック()
    '変数設定
    Dim seiseki_tourokuSheet As Worksheet
    Dim seiseki_kakuninnSheet As Worksheet
    Dim db_seiseki_infoSheet As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim i As Long
    
    'シート設定
    Set seiseki_tourokuSheet = ThisWorkbook.Sheets("成績登録用")
    Set seiseki_kakuninnSheet = ThisWorkbook.Sheets("成績確認用")
    Set db_seiseki_infoSheet = ThisWorkbook.Sheets("DB（成績情報）")
    
    '開始行と終了行の取得
    startRow = 12
    lastRow = seiseki_tourokuSheet.Cells(seiseki_tourokuSheet.Rows.Count, "A").End(xlUp).Row
    
    '成績情報をDB（成績情報）シートに登録
    For i = startRow To lastRow
        Dim yearValue As String
        Dim classValue As String
        Dim numberValue As String
        Dim nameValue As String
        Dim score As Double
        Dim testName As String
        Dim fullScore As Double
    
        '生徒情報の取得
        yearValue = seiseki_tourokuSheet.Cells(i, "A").Value
        classValue = seiseki_tourokuSheet.Cells(i, "B").Value
        numberValue = seiseki_tourokuSheet.Cells(i, "C").Value
        nameValue = seiseki_tourokuSheet.Cells(i, "D").Value
      
        '成績情報の取得
        score = seiseki_tourokuSheet.Cells(i, "E").Value
        
        'テスト名と満点の取得
        testName = seiseki_tourokuSheet.Cells(5, "A").Value
        fullScore = seiseki_tourokuSheet.Cells(8, "A").Value
        
        'DB（成績情報）シートに情報を追加
        Dim db_seiseki_infoLastRow As Long
        db_seiseki_infoLastRow = db_seiseki_infoSheet.Cells(db_seiseki_infoSheet.Rows.Count, "A").End(xlUp).Row + 1
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "A").Value = testName
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "B").Value = fullScore
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "C").Value = yearValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "D").Value = classValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "E").Value = numberValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "F").Value = nameValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "G").Value = score
        ' 登録時刻を記録
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "H").Value = Now()
        
        '格子状の線も追加
        db_seiseki_infoSheet.Range("A" & db_seiseki_infoLastRow & ":H" & db_seiseki_infoLastRow).Borders.LineStyle = xlContinuous
    Next i
    
        '成績算出のプロシージャ呼出
        
        
        'メッセージ表示
        MsgBox "成績情報を登録しました。成績確認に進んでください。", vbInformation
        
        '登録シートをクリア
        seiseki_tourokuSheet.Range("A5:E5").ClearContents
        seiseki_tourokuSheet.Range("A8:C8").ClearContents
        seiseki_tourokuSheet.Range("E" & startRow & ":E" & lastRow).ClearContents
End Sub
