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

Sub ���я��o�^�{�^���N���b�N()
    '�ϐ��ݒ�
    Dim seiseki_tourokuSheet As Worksheet
    Dim seiseki_kakuninnSheet As Worksheet
    Dim db_seiseki_infoSheet As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim i As Long
    
    '�V�[�g�ݒ�
    Set seiseki_tourokuSheet = ThisWorkbook.Sheets("���ѓo�^�p")
    Set seiseki_kakuninnSheet = ThisWorkbook.Sheets("���ъm�F�p")
    Set db_seiseki_infoSheet = ThisWorkbook.Sheets("DB�i���я��j")
    
    '�J�n�s�ƏI���s�̎擾
    startRow = 12
    lastRow = seiseki_tourokuSheet.Cells(seiseki_tourokuSheet.Rows.Count, "A").End(xlUp).Row
    
    '���я���DB�i���я��j�V�[�g�ɓo�^
    For i = startRow To lastRow
        Dim yearValue As String
        Dim classValue As String
        Dim numberValue As String
        Dim nameValue As String
        Dim score As Double
        Dim testName As String
        Dim fullScore As Double
    
        '���k���̎擾
        yearValue = seiseki_tourokuSheet.Cells(i, "A").Value
        classValue = seiseki_tourokuSheet.Cells(i, "B").Value
        numberValue = seiseki_tourokuSheet.Cells(i, "C").Value
        nameValue = seiseki_tourokuSheet.Cells(i, "D").Value
      
        '���я��̎擾
        score = seiseki_tourokuSheet.Cells(i, "E").Value
        
        '�e�X�g���Ɩ��_�̎擾
        testName = seiseki_tourokuSheet.Cells(5, "A").Value
        fullScore = seiseki_tourokuSheet.Cells(8, "A").Value
        
        'DB�i���я��j�V�[�g�ɏ���ǉ�
        Dim db_seiseki_infoLastRow As Long
        db_seiseki_infoLastRow = db_seiseki_infoSheet.Cells(db_seiseki_infoSheet.Rows.Count, "A").End(xlUp).Row + 1
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "A").Value = testName
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "B").Value = fullScore
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "C").Value = yearValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "D").Value = classValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "E").Value = numberValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "F").Value = nameValue
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "G").Value = score
        ' �o�^�������L�^
        db_seiseki_infoSheet.Cells(db_seiseki_infoLastRow, "H").Value = Now()
        
        '�i�q��̐����ǉ�
        db_seiseki_infoSheet.Range("A" & db_seiseki_infoLastRow & ":H" & db_seiseki_infoLastRow).Borders.LineStyle = xlContinuous
    Next i
    
        '���юZ�o�̃v���V�[�W���ďo
        
        
        '���b�Z�[�W�\��
        MsgBox "���я���o�^���܂����B���ъm�F�ɐi��ł��������B", vbInformation
        
        '�o�^�V�[�g���N���A
        seiseki_tourokuSheet.Range("A5:E5").ClearContents
        seiseki_tourokuSheet.Range("A8:C8").ClearContents
        seiseki_tourokuSheet.Range("E" & startRow & ":E" & lastRow).ClearContents
End Sub
