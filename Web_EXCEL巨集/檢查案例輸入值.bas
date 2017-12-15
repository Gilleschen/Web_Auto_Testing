Attribute VB_Name = "檢查案例輸入值"
Function CheckValue()
    Dim sheetname As String
    Dim xpath, id As String
    xpath = "xpath": id = "id"
    i = 0
    Do
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
            'If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
                sheetname = ThisWorkbook.Sheets(i + 1).Name
                j = 1
                Do
                
                    Select Case Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A")
                    
                    Case "CaseName"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行缺少CaseName", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
            
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                         CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "Byid_Click"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                    
                    Case "ByXpath_Click"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        'CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        
                        
                    Case "Byid_Clear"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                    
                    Case "ByXpath_Clear"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        'CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function

                    Case "Byid_SendKey"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkSendKey(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
                                            
                    Case "ByXpath_SendKey"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkSendKey(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "D")
                        If CheckValue = False Then Exit Function
                        'CheckValue = checkXpath(sheetname, i, j)
                        
                    
                    Case "Byid_Scroll"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkScroll(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_Scroll"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkScroll(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                       'CheckValue = checkXpath(sheetname, i, j)
                        'If CheckValue = False Then Exit Function
                        
                    Case "Byid_invisibility"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_invisibility"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkClick(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        'CheckValue = checkXpath(sheetname, i, j)
                        If CheckValue = False Then Exit Function
                        
                    Case "Launch"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
    
                            x = MsgBox(sheetname & "中，第" & j & "列僅能填入Launch", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                    
                    Case "Quit"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            
                            x = MsgBox(sheetname & "中，第" & j & "列僅能填入Quit", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Back"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            
                            x = MsgBox(sheetname & "中，第" & j & "列僅能填入Back", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Next"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            
                            x = MsgBox(sheetname & "中，第" & j & "列僅能填入Next", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Refresh"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            
                            x = MsgBox(sheetname & "中，第" & j & "列僅能填入Refresh", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                    Case "Goto"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
                            
                            x = MsgBox(sheetname & "中，第" & j & "列請填入網址 (如 https://www.google.com.tw)", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0): CheckValue = False: Exit Function
                        
                        ElseIf Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                        
                            If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 4) <> "http" Or Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 4) <> "http" Then
                                
                                x = MsgBox(sheetname & "中，第" & j & "列網址前面請加入https://或http://", 0 + 16, "Error"): Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0): CheckValue = False: Exit Function
                            Else
                            
                                Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                                CheckValue = True
                                    
                            End If
                       
                        End If
                        
                    Case "Byid_VerifyText"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        'If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
                        CheckValue = checkResult(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ByXpath_VerifyText"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        'If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
                        CheckValue = checkResult(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        'CheckValue = checkXpath(sheetname, i, j)

                    Case "Byid_Wait"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkWait(sheetname, i, j, id)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                    
                    Case "ByXpath_Wait"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        CheckValue = checkWait(sheetname, i, j, xpath)
                        If CheckValue = False Then Exit Function
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        'CheckValue = checkXpath(sheetname, i, j)
                       
                    Case "Sleep"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
                            x = MsgBox(sheetname & "中，第" & j & "行缺少秒數", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                            
                        ElseIf IsNumeric(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")) = False Then
                            x = MsgBox(sheetname & "中，第" & j & "行請輸入數值", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        
                        Else
                            If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Value) <> "String" Then
                               Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B")
                            End If
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If
                        
                        CheckValue = checkExcessData(sheetname, i, j, "C")
                        If CheckValue = False Then Exit Function
                        
                    Case "ScreenShot"
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(0, 0, 0)
                        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") <> "" Then
                            x = MsgBox(sheetname & "中，第" & j & "僅能填入ScreenShot", 0 + 16, "Error")
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
                            CheckValue = False
                            Exit Function
                        Else
                            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
                            CheckValue = True
                        End If

                    
                    Case Else
                        x = MsgBox(sheetname & "中，第" & j & "行語法有誤，" & "無" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Value & " 語法", 0 + 16, "Error")
                        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A").Font.Color = RGB(255, 0, 0)
                        CheckValue = False
                        Exit Function
                    End Select
                    
                    
                j = j + 1
                Loop Until Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "A") = ""
            
           ' End If
    
            
        End If
        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    CheckValue2 = Delete_All_Blank_Cells
End Function

Function checkExcessData(sheetname, i, j, col) '檢查所有指令最後一欄是否為空白或無資料

    If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, col) <> "" Then
                                
        x = MsgBox(sheetname & "中，第" & j & "列第3欄請保持空白", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, col).Interior.Color = RGB(255, 0, 0)
        checkExcessData = False:
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, col).Interior.Pattern = xlNone
        checkExcessData = True
    End If

End Function

Function checkXpath(sheetname, i, j)
    
    If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) <> "//*[@" And Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 6) <> "(//*[@" Then
        x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
        checkXpath = False
        Exit Function
    'ElseIf Right(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 1) <> "]" Then
        'x = MsgBox(sheetname & "中，第" & j & "行xpath有誤", 0 + 16, "Error")
        'Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
   
    Else
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "E").Interior.Pattern = xlNone
        Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Font.Color = RGB(0, 0, 0)
        checkXpath = True
    End If
    
End Function


Function checkClick(sheetname, i, j, status)
    
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkClick = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkClick = True
        End If
        
    Else
    
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkClick = False
            Exit Function
        ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkClick = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkClick = True
        End If
        
    End If

End Function
Function checkScroll(sheetname, i, j, status)
    
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkScroll = False
            Exit Function
            
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkScroll = True
        End If
        
    Else
    
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkScroll = False
            Exit Function
        ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid_Scroll，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkScroll = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkScroll = True
        End If
        
    End If

End Function

Function checkWait(sheetname, i, j, status)
    
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkWait = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkWait = True
        End If
        
    Else
    
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkWait = False
            Exit Function
        ElseIf Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid_Wait，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkWait = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkWait = True
        End If

        
    End If

End Function

Function checkResult(sheetname, i, j, status)
    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkResult = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkResult = True
        End If
    
    Else
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkResult = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkResult = True
        End If
        
        If Left(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B"), 5) = "//*[@" Then
            x = MsgBox(sheetname & "中，第" & j & "使用Byid_VerifyText，卻輸入Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkResult = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkResult = True
            
        End If
    
    End If
End Function

Function checkSendKey(sheetname, i, j, status)

    If status = "xpath" Then
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件Xpath", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkSendKey = True
        End If
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少輸入值", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
        Else
            
            If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Value) <> "String" Then
                Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C")
            End If
        
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSendKey = True
        End If
    Else
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少元件id", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
        Else
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "B").Interior.Pattern = xlNone
            checkSendKey = True
        End If
        
        If Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "" Then
            x = MsgBox(sheetname & "中，第" & j & "行缺少輸入值", 0 + 16, "Error")
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Color = RGB(255, 0, 0)
            checkSendKey = False
            Exit Function
        Else
            If TypeName(Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Value) <> "String" Then
                Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C") = "'" & Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C")
            End If
            Sheets(ThisWorkbook.Sheets(i + 1).Name).Cells(j, "C").Interior.Pattern = xlNone
            checkSendKey = True
        End If
                
    End If
    
End Function
Function Clear_Hidekeyboard_LaunchAPP_QuitAPP()
    Application.ScreenUpdating = False
    Dim sheetname As String
    
    i = 0
    Do
        
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
            If ThisWorkbook.Sheets(i + 1).Visible = True Then
                        
                'sheetname = ThisWorkbook.Sheets(i + 1).Name
                'Sheets(sheetname).Select
                ThisWorkbook.Sheets(i + 1).Select
                j = 1
                Do
                    If ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "ScreenShot" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "ResetAPP" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Power" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Home" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Back" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "QuitAPP" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "LaunchAPP" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "HideKeyboard" Or ThisWorkbook.Sheets(i + 1).Cells(j, "A").Value = "Menu" Then
                        For k = 1 To 5
                            ThisWorkbook.Sheets(i + 1).Cells(j, "B").Select
                            Selection.Delete Shift:=xlToLeft
                        Next k
                    End If
                    
                
                    j = j + 1
                Loop Until ThisWorkbook.Sheets(i + 1).Cells(j, "A") = ""
    
            End If
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    Sheets("APP&Device").Select
End Function

Function Delete_All_Blank_Cells()
    Application.ScreenUpdating = False
    Dim sheetname As String
    
    i = 0
    Do
        
        If Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
        
            If ThisWorkbook.Sheets(i + 1).Visible = True Then
                 
                ThisWorkbook.Sheets(i + 1).Select
                j = 1
                Do
                    k = 1
                    Do While ThisWorkbook.Sheets(i + 1).Cells(j, k) <> ""
                        k = k + 1
                    Loop
                       
                    For w = 1 To 5
                        ThisWorkbook.Sheets(i + 1).Cells(j, k).Select
                        Selection.Delete Shift:=xlToLeft
                    Next w
                       
                       
                       
                j = j + 1
                Loop Until ThisWorkbook.Sheets(i + 1).Cells(j, "A") = ""
        
            End If
        End If

        i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
    Sheets("Web_Infor").Select
End Function

