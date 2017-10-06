Attribute VB_Name = "檢查案例語法"
Function CheckCommand()
    Dim sheetname As String
    Dim CaseName, LaunchAPP, Byid_Result, ByXpath_Result, QuitAPP As Integer
    
    i = 0
    Do
        
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" And ThisWorkbook.Sheets(i + 1).Visible = True Then
            'If ThisWorkbook.Sheets(i + 1).Visible = True Then
                CaseName = 0: LaunchAPP = 0: Byid_Result = 0: ByXpath_Result = 0: QuitAPP = 0
                sheetname = ThisWorkbook.Sheets(i + 1).Name
                j = 1
                Do
                
                    Select Case Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A")
                    
                    Case "CaseName"
                    
                        CaseName = CaseName + 1
                    
                    Case "Launch"
                    
                        LaunchAPP = LaunchAPP + 1
                    
                    Case "Quit"
                    
                        QuitAPP = QuitAPP + 1
                    
                    'Case "Byid_Result"
    
                        'Byid_Result = Byid_Result + 1
                        
                    'Case "ByXpath_Result"
                        
                        'ByXpath_Result = ByXpath_Result + 1
    
                    End Select
                    
                j = j + 1
                Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
                
                If LaunchAPP <> CaseName Then
                    x = MsgBox(sheetname & "中缺少LaunchAPP或CaseName", 0 + 16, "Error")
                    CheckCommand = False
                    Exit Function
                Else
                    CheckCommand = True
                End If
 
                If QuitAPP <> CaseName Then
                    x = MsgBox(sheetname & "中缺少QuitAPP或CaseName", 0 + 16, "Error")
                    CheckCommand = False
                    Exit Function
                Else
                    CheckCommand = True
                End If

                'If Byid_Result <> CaseName Or ByXpath_Result <> CaseName Then x = MsgBox(sheetname & "中缺少Byid_Result或CaseName", 0 + 16, "Error")
              
            'End If
        
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
End Function


