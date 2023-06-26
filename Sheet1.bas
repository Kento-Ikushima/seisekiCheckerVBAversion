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

Sub ���k���o�^�{�^���N���b�N()
    Dim seiseki_tourokuSheet As Worksheet
    Dim seiseki_kakuninnSheet As Worksheet
    Dim regSheet As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim i As Long
    
    ' �V�[�g�̐ݒ�
    Set seiseki_tourokuSheet = ThisWorkbook.Sheets("���ѓo�^�p")
    Set seiseki_kakuninnSheet = ThisWorkbook.Sheets("���ъm�F�p")
    Set regSheet = ThisWorkbook.Sheets("���k�o�^�p")
    
    ' �J�n�s�ƏI���s�̎擾
    startRow = 5
    lastRow = regSheet.Cells(regSheet.Rows.Count, "A").End(xlUp).Row
    
    ' ���k�����Q�̃V�[�g�Ɉꊇ�o�^
    For i = startRow To lastRow
        Dim yearValue As String
        Dim classValue As String
        Dim numberValue As String
        Dim nameValue As String
        
        ' ���͒l�̎擾
        yearValue = regSheet.Cells(i, "A").Value
        classValue = regSheet.Cells(i, "B").Value
        numberValue = regSheet.Cells(i, "C").Value
        nameValue = regSheet.Cells(i, "D").Value
        
        ' ���ѓo�^�p�V�[�g�ɏ���ǉ�
        Dim seiseki_tourokuLastRow As Long
        seiseki_tourokuLastRow = seiseki_tourokuSheet.Cells(seiseki_tourokuSheet.Rows.Count, "A").End(xlUp).Row + 1
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "A").Value = yearValue
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "B").Value = classValue
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "C").Value = numberValue
        seiseki_tourokuSheet.Cells(seiseki_tourokuLastRow, "D").Value = nameValue
        ' �i�q��̐���ǉ�
        seiseki_tourokuSheet.Range("A" & seiseki_tourokuLastRow & ":E" & seiseki_tourokuLastRow).Borders.LineStyle = xlContinuous
        
        ' ���ъm�F�p�V�[�g�ɏ���ǉ�
        Dim seiseki_kakuninnLastRow As Long
        seiseki_kakuninnLastRow = seiseki_kakuninnSheet.Cells(seiseki_kakuninnSheet.Rows.Count, "A").End(xlUp).Row + 1
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "A").Value = yearValue
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "B").Value = classValue
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "C").Value = numberValue
        seiseki_kakuninnSheet.Cells(seiseki_kakuninnLastRow, "D").Value = nameValue
        
        ' �i�q��̐���ǉ�
        seiseki_kakuninnSheet.Range("A" & seiseki_kakuninnLastRow & ":E" & seiseki_kakuninnLastRow).Borders.LineStyle = xlContinuous
    Next i
    
    ' ���b�Z�[�W�\��
    MsgBox "���k��o�^���܂����B���ѓo�^�ɐi��ł��������B", vbInformation
    
    ' �o�^�p�V�[�g���N���A
    regSheet.Range("A5:D" & lastRow).ClearContents
End Sub

