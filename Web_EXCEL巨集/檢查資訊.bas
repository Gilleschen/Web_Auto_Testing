Attribute VB_Name = "檢查資訊"
Function CheckAPPandDevice()
    Dim sheetname As String
    Dim scriptnumber, result As Integer
    'Dim BrowserList(5) As String
    
    BrowserList = Array("chrome", "firefox", "internet explorer", "safari", "opera")
    Application.ScreenUpdating = False
    
    i = 1 '檢查各欄位是否填入資料
    Do
        
        If Sheets("Web_Infor").Cells(2, i) = "" Then
            
           x = MsgBox("請填入" & Sheets("Web_Infor").Cells(1, i), 0 + 16, "Error")
           Sheets("Web_Infor").Cells(2, i).Interior.Color = RGB(255, 0, 0)
           CheckAPPandDevice = False
           Exit Function
        Else
        
            Sheets("Web_Infor").Cells(2, i).Interior.Pattern = xlNone
            CheckAPPandDevice = True
                
        End If
    
        i = i + 1
    Loop Until Sheets("Web_Infor").Cells(1, i) = ""
    

    '計算ScriptName欄位下，待測試的腳本數量
    j = 2: scriptnumber = 0
    Do
        scriptnumber = scriptnumber + 1
    j = j + 1
    Loop Until Sheets("Web_Infor").Cells(j, "D") = ""
    
    '根據統計的測試腳本數量，重新定義scriptarray為陣列
    ReDim scriptarray(scriptnumber - 1) As String
    
    '待測試的腳本名稱加入scriptarray陣列
    j = 2: x = 0
    Do
        scriptarray(x) = Sheets("Web_Infor").Cells(j, "D")
    j = j + 1: x = x + 1
    Loop Until Sheets("Web_Infor").Cells(j, "D") = ""
    
    
    i = 0
    Do
        j = 0: result = 0
        Do
            sheetname = ThisWorkbook.Sheets(j + 1).Name
            If scriptarray(i) <> sheetname Then result = result + 1
    
            j = j + 1
        Loop Until j = ThisWorkbook.Sheets.Count
        If result = ThisWorkbook.Sheets.Count Then
            y = MsgBox("找不到" & scriptarray(i) & "工作表", 0 + 16, "Error")
            CheckAPPandDevice = False
            Exit Function
        Else
            CheckAPPandDevice = True
        End If
        i = i + 1
    Loop Until i = UBound(scriptarray) - LBound(scriptarray) + 1
    
    i = 2
    Do
        If Right(Sheets("Web_Infor").Cells(i, "D"), 11) <> "_TestScript" Then
            
            y = MsgBox("ScriptName欄位請填入以_TestScript為結尾之工作表(大小寫有分)", 0 + 16, "Error")
            Sheets("Web_Infor").Cells(i, "D").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            Sheets("Web_Infor").Cells(i, "D").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
    
    i = i + 1
    Loop Until Sheets("Web_Infor").Cells(i, "D") = ""
    
    '確認Browser欄位內容
    i = 2
    Do
        j = 0: result = 0
        Do
            If Sheets("Web_Infor").Cells(i, "A") <> BrowserList(j) Then result = result + 1
        j = j + 1
        Loop Until j = UBound(BrowserList) - LBound(BrowserList) + 1
        
        If result = UBound(BrowserList) - LBound(BrowserList) + 1 Then
            y = MsgBox(Sheets("Web_Infor").Cells(i, "A") & "格式錯誤" & vbNewLine & "請輸入：chrome, firefox, internet explorer, safari, opera" & vbNewLine & "(全英文小寫)", 0 + 16, "Error")
            Sheets("Web_Infor").Cells(i, "A").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            Sheets("Web_Infor").Cells(i, "A").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
        
    i = i + 1
    Loop Until Sheets("Web_Infor").Cells(i, "A") = ""
    
    
    
    '確認BrowserDriverPath
    j = 2
    Do
        If Sheets("Web_Infor").Cells(j, "B") = "" Then
            y = MsgBox("請填入" & Sheets("Web_Infor").Cells(j, "A").Value & "之BrowserDriverPath", 0 + 16, "Error")
            Sheets("Web_Infor").Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            If Dir(CStr(Sheets("Web_Infor").Cells(j, "B"))) = "" Then
                x = MsgBox("找不到" & Sheets("Web_Infor").Cells(j, "B"), 0 + 16, "Error")
                Sheets("Web_Infor").Cells(j, "B").Font.Color = RGB(255, 0, 0)
                CheckAPPandDevice = False
                Exit Function
            Else
                Sheets("Web_Infor").Cells(j, "B").Font.Color = RGB(0, 0, 0)
                CheckAPPandDevice = True
            End If
        End If
    
    j = j + 1
    Loop Until Sheets("Web_Infor").Cells(j, "A") = ""
    
    '檢查SeleniumServerJarPath
    If Sheets("Web_Infor").Cells(2, "F") = "" Then
    
        x = MsgBox("請填入SeleniumServerJarPath檔路徑" & vbNewLine & "例如：C:\Users\Desktop\檔名.jar", 0 + 16, "Error")
        CheckAPPandDevice = False
        Exit Function
    Else
        If Dir(CStr(Sheets("Web_Infor").Cells(2, "F"))) = "" Then
        
            x = MsgBox("找不到" & Sheets("Web_Infor").Cells(2, "F"), 0 + 16, "Error")
            Sheets("Web_Infor").Cells(2, "F").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            Sheets("Web_Infor").Cells(2, "F").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
    
    End If
    
    '檢查JarPath
    If Sheets("Web_Infor").Cells(2, "E") = "" Then
    
        x = MsgBox("請填入JarPath檔路徑" & vbNewLine & "例如：C:\Users\Desktop\檔名.jar", 0 + 16, "Error")
        CheckAPPandDevice = False
        Exit Function
    Else
    
        If Dir(CStr(Sheets("Web_Infor").Cells(2, "E"))) = "" Then
        
            x = MsgBox("找不到" & Sheets("Web_Infor").Cells(2, "E"), 0 + 16, "Error")
            Sheets("Web_Infor").Cells(2, "E").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
            
        Else
            Sheets("Web_Infor").Cells(2, "E").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
    End If
    
    
End Function
