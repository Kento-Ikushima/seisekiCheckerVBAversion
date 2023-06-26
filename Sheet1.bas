VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub 生徒情報登録ボタンクリック()
    Dim seiseki_tourokuSheet As Worksheet
    Dim seiseki_kakuninnSheet As Worksheet
    Dim regSheet As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim i As Long
    
    ' シートの設定
    Set seiseki_tourokuSheet = ThisWorkbook.Sheets("成績登録用")
    Set seiseki_kakuninnSheet = ThisWorkbook.Sheets("成績確認用")
    Set regSheet = ThisWorkbook.Sheets("生徒登録用")
    
    ' 開始行と終了行の取得
    startRow = 5
    lastRow = regSheet.Cells(regSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 生徒情報を２つのシートに一括登録
    For i = startRow To lastRow
        Dim yearValue As String
        Dim classValue As String
        Dim numberValue As String
        Dim nameValue As String
        
        ' 入力値の取得
        yearValue = regSheet.Cells(i, "A").Value
        classValue = regSheet.Cells(i, "B").Value
        numberValue = regSheet.Cells(i, "C").Value
        nameValue = regSheet.Cells(i, "D").Value
        
        ' 成績登録用シートに情報を追加
        Dim seiseki_tourokuLastRow As Long
        seiseki_tourokuLastRow = seiseki_tourokuSheet.Cells(seiseki_tourokuSheet.Rows.Count, "A").End(xlUp).Row + 1
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "A").Value = yearValue
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "B").Value = classValue
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "C").Value = numberValue
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "D").Value = nameValue
        ' 格子状の線を追加
        seiseki_tourokuSheet.Range("A" & seiseki_tourokuLastRow & ":E" & seiseki_tourokuLastRow).Borders.LineStyle = xlContinuous
        
        ' 成績確認用シートに情報を追加
        Dim seiseki_kakuninnLastRow As Long
        seiseki_kakuninnLastRow = seiseki_kakuninnSheet.Cells(seiseki_kakuninnSheet.Rows.Count, "A").End(xlUp).Row + 1
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "A").Value = yearValue
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "B").Value = classValue
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "C").Value = numberValue
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "D").Value = nameValue
        
        ' 格子状の線を追加
        seiseki_kakuninnSheet.Range("A" & seiseki_kakuninnLastRow & ":E" & seiseki_kakuninnLastRow).Borders.LineStyle = xlContinuous
    Next i
    
    ' メッセージ表示
    MsgBox "生徒を登録しました。成績登録に進んでください。", vbInformation
    
    ' 登録用シートをクリア
    regSheet.Range("A5:D" & lastRow).ClearContents
End Sub

