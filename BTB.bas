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
ans = MsgBox("BTBカラーミーショップからのデータを処理しますか？" & vbCrLf & "【　" & i & "　件】", vbOKCancel, "【確認】")

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
      '輸送サービスの行削
      
      '↓重複IDの行削除
      
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
       '↑ここまで
       
       '行数カウント　ヘッダーこみ
       i = 1
       Do While wb1.ActiveSheet.Cells(i, 1) <> ""
          i = i + 1
       Loop
       
       Workbooks.Open filename:= _
               "\\hdd-tps2\share\TPS部\毎日出荷フォルダー\999 BTB\カラーミー荷札出力まとめ.xlsx"
     Set wb2 = Workbooks("カラーミー荷札出力まとめ.xlsx")
     wb2.Sheets("取込").Range("a1:z1000").ClearContents
     wb1.ActiveSheet.Cells.Copy Workbooks("カラーミー荷札出力まとめ.xlsx").Sheets("取込").Cells   'データをコピー
     wb1.Close SaveChanges:=False
     '－－－－－－ここまで　カラミ―CSVから　行を削除したりしてはりつけ
     '－－－－－－ここから　まとめエクセルでの貼付や累積やCSVに変換
     
     wb2.Sheets("csv").Range("a1:q1000").ClearContents
     wb2.Sheets("コピーして").Range("a2:Q2").Resize(i - 2).Copy
     wb2.Sheets("csv").Range("A1").PasteSpecial Paste:=xlPasteValues
     
     j = 1
     Do While wb2.Sheets("まとめ").Cells(j, 1) <> ""
          j = j + 1
     Loop
     
     wb2.Sheets("まとめ").Cells(j, 1).PasteSpecial Paste:=xlPasteValues
     
     wb2.Save
            
            'CSVに　保存は飛電
            Sheets("csv").Select
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs filename:= _
                "\\SAGAWA-HP\Users\Public\Documents\伝発用\btb.csv", FileFormat:= _
                xlCSV, Local:=True
                 Application.DisplayAlerts = True
     
          Workbooks.Open filename:= _
               "\\hdd-tps2\share\TPS部\毎日出荷フォルダー\999 BTB\カラーミー荷札出力まとめ.xlsx"
               
               Workbooks("btb.csv").Close SaveChanges:=False
     MsgBox "処理終了☆♪" & i - 2 & "件出力"
Else
     MsgBox "キャンセル"
End If

End Sub

