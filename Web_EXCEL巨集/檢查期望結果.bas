Attribute VB_Name = "�ˬd���浲�G"
Sub CheckExpectResult()
    Dim result As String
    Dim x As Integer
    i = 0
    Do
        If Right(ThisWorkbook.Sheets(i + 1).Name, 9) = "_TestCase" And ThisWorkbook.Sheets(i + 1).Visible = True Then
        
            'MsgBox (ThisWorkbook.Sheets(i + 1).Name)
            
            j = 1: x = 0
            Do
                If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = "CaseName" Then
                    
                    x = x + 1
                
                End If
            
                
                j = j + 1
            Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
            
            ReDim casename(x - 1)
            
            j = 1: x = 0
            
            Do
                If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = "CaseName" Then
                    
                    casename(x) = Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")
                    x = x + 1
                    
                End If
            
                j = j + 1
            Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
        
        
        
            k = 0
            Do
                j = 2
                Do
                    
                    If casename(k) = Sheets("ExpectResult").Cells(j, "A") Then
                        result = "Pass"
                        Exit Do
                    End If
                    
                    j = j + 1
                Loop Until Sheets("ExpectResult").Cells(j, "A") = ""
                
                If result <> "Pass" Then x = MsgBox(casename(k) + "�����浲�G�����g�JExpectResult", 0 + 16, "Error")
                
                result = ""
                
                k = k + 1
            Loop Until k = UBound(casename) - LBound(casename) + 1
        
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
End Sub
