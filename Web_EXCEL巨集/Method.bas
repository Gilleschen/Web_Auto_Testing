Attribute VB_Name = "Method"
Sub Classification_TestCase()
    Dim row As String
    Dim color As Integer
    color = 1
    
    Application.ScreenUpdating = False
   
    i = 0
    Do
        start_count = 1
        Count = 1
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
        
            sheetname = ThisWorkbook.Sheets(i + 1).Name
            Sheets(sheetname).Select
            j = 1
            
            Do
               
                Do
                
                    Count = Count + 1
            
                Loop Until Sheets(sheetname).Cells(Count, "A") = "CaseName" Or Sheets(sheetname).Cells(Count, "A") = ""
                
                row = start_count & ":" & Count - 1
                start_count = Count
                
                color = color * (-1)
                
                Rows(row).Select
                
                If color < 0 Then
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.8
                    .PatternTintAndShade = 0
                End With
                Else
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.8
                    .PatternTintAndShade = 0
                End With
                End If
            j = start_count
            Loop Until Sheets(sheetname).Cells(j, "A") = ""
            
        End If
    i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
End Sub
