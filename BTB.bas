Attribute VB_Name = "BTB"
Option Explicit

Sub BTB()

Dim i As Integer
Dim j As Integer
Dim ans As Variant
Dim wb1 As Workbook
Dim wb2 As Workbook

Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic

Set wb1 = ActiveWorkbook
i = 2
Do While Cells(i, 1) <> ""
     i = i + 1
Loop
i = i - 2
ans = MsgBox("BTB�J���[�~�[�V���b�v����̃f�[�^���������܂����H" & vbCrLf & "�y�@" & i & "�@���z", vbOKCancel, "�y�m�F�z")

If ans = vbOK Then
     With wb1.ActiveSheet
          j = i + 1
          Do While j > 1
               Select Case .Cells(j, 58)
                    Case 75151795
                         .Rows(j).Delete
                    Case 77692678
                         .Rows(j).Delete
                    Case 77827848
                         .Rows(j).Delete
                    Case 77829120
                         .Rows(j).Delete
                    Case 72980433
                         .Rows(j).Delete
                    Case 72980741
                         .Rows(j).Delete
                    Case 126134693
                         .Rows(j).Delete
                    Case 126136429
                         .Rows(j).Delete
                    Case 126138540
                         .Rows(j).Delete
                    Case 126002193
                         .Rows(j).Delete
                    Case 126141583
                         .Rows(j).Delete
                    Case 126143475
                         .Rows(j).Delete
                    Case 126144097
                         .Rows(j).Delete
                    Case 126144187
                         .Rows(j).Delete
                    Case 126144401
                         .Rows(j).Delete
                    Case 126144600
                         .Rows(j).Delete
                    Case 126165011
                         .Rows(j).Delete
                         
               End Select
                         
               j = j - 1
          Loop
      '�A���T�[�r�X�̍s��
      
      '���d��ID�̍s�폜
      
          i = i + 1
          Do While i > 1
              If .Cells(i, 1) = .Cells(i - 1, 1) Then
                   .Rows(i).Delete
              End If
              i = i - 1
          Loop
   
          .Columns("BC:BK").Delete
          .Columns("AS:BA").Delete
          .Columns("y:ah").Delete
          .Columns("p:v").Delete
          .Columns("k").Delete
          .Columns("c").Delete
         
          i = 2
          Do While .Cells(i, 1) <> ""
               .Cells(i, 26) = Replace(.Cells(i, 26), vbLf, "")
               i = i + 1
          Loop
      End With
       '�������܂�
       
       '�s���J�E���g�@�w�b�_�[����
       i = 1
       Do While wb1.ActiveSheet.Cells(i, 1) <> ""
          i = i + 1
       Loop
       
       Workbooks.Open filename:= _
               "\\hdd-tps2\share\TPS��\�����o�׃t�H���_�[\999 BTB\�J���[�~�[�׎D�o�͂܂Ƃ�.xlsx"
     Set wb2 = Workbooks("�J���[�~�[�׎D�o�͂܂Ƃ�.xlsx")
     wb2.Sheets("�捞").Range("a1:z1000").ClearContents
     wb1.ActiveSheet.Cells.Copy Workbooks("�J���[�~�[�׎D�o�͂܂Ƃ�.xlsx").Sheets("�捞").Cells   '�f�[�^���R�s�[
     wb1.Close SaveChanges:=False
     '�|�|�|�|�|�|�����܂Ł@�J���~�\CSV����@�s���폜�����肵�Ă͂��
     '�|�|�|�|�|�|��������@�܂Ƃ߃G�N�Z���ł̓\�t��ݐς�CSV�ɕϊ�
     
     wb2.Sheets("csv").Range("a1:q1000").ClearContents
     wb2.Sheets("�R�s�[����").Range("a2:Q2").Resize(i - 2).Copy
     wb2.Sheets("csv").Range("A1").PasteSpecial Paste:=xlPasteValues
     
     j = 1
     Do While wb2.Sheets("�܂Ƃ�").Cells(j, 1) <> ""
          j = j + 1
     Loop
     
     wb2.Sheets("�܂Ƃ�").Cells(j, 1).PasteSpecial Paste:=xlPasteValues
     
     wb2.Save
            
            'CSV�Ɂ@�ۑ��͔�d
            Sheets("csv").Select
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs filename:= _
                "\\SAGAWA-HP\Users\Public\Documents\�`���p\btb.csv", FileFormat:= _
                xlCSV, Local:=True
                 Application.DisplayAlerts = True
     
          Workbooks.Open filename:= _
               "\\hdd-tps2\share\TPS��\�����o�׃t�H���_�[\999 BTB\�J���[�~�[�׎D�o�͂܂Ƃ�.xlsx"
               
               Workbooks("btb.csv").Close SaveChanges:=False
     MsgBox "�����I������" & i - 2 & "���o��"
Else
     MsgBox "�L�����Z��"
End If

End Sub

