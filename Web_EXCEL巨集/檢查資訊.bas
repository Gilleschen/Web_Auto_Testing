Attribute VB_Name = "�ˬd��T"
Function CheckAPPandDevice()
    Dim sheetname As String
    Dim scriptnumber, result As Integer
    'Dim BrowserList(5) As String
    
    BrowserList = Array("chrome", "firefox", "internet explorer", "safari", "opera")
    Application.ScreenUpdating = False
    
    i = 1 '�ˬd�U���O�_��J���
    Do
        
        If Sheets("Web_Infor").Cells(2, i) = "" Then
            
           x = MsgBox("�ж�J" & Sheets("Web_Infor").Cells(1, i), 0 + 16, "Error")
           Sheets("Web_Infor").Cells(2, i).Interior.Color = RGB(255, 0, 0)
           CheckAPPandDevice = False
           Exit Function
        Else
        
            Sheets("Web_Infor").Cells(2, i).Interior.Pattern = xlNone
            CheckAPPandDevice = True
                
        End If
    
        i = i + 1
    Loop Until Sheets("Web_Infor").Cells(1, i) = ""
    

    '�p��ScriptName���U�A�ݴ��ժ��}���ƶq
    j = 2: scriptnumber = 0
    Do
        scriptnumber = scriptnumber + 1
    j = j + 1
    Loop Until Sheets("Web_Infor").Cells(j, "D") = ""
    
    '�ھڲέp�����ո}���ƶq�A���s�w�qscriptarray���}�C
    ReDim scriptarray(scriptnumber - 1) As String
    
    '�ݴ��ժ��}���W�٥[�Jscriptarray�}�C
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
            y = MsgBox("�䤣��" & scriptarray(i) & "�u�@��", 0 + 16, "Error")
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
            
            y = MsgBox("ScriptName���ж�J�H_TestScript���������u�@��(�j�p�g����)", 0 + 16, "Error")
            Sheets("Web_Infor").Cells(i, "D").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            Sheets("Web_Infor").Cells(i, "D").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
    
    i = i + 1
    Loop Until Sheets("Web_Infor").Cells(i, "D") = ""
    
    '�T�{Browser��줺�e
    i = 2
    Do
        j = 0: result = 0
        Do
            If Sheets("Web_Infor").Cells(i, "A") <> BrowserList(j) Then result = result + 1
        j = j + 1
        Loop Until j = UBound(BrowserList) - LBound(BrowserList) + 1
        
        If result = UBound(BrowserList) - LBound(BrowserList) + 1 Then
            y = MsgBox(Sheets("Web_Infor").Cells(i, "A") & "�榡���~" & vbNewLine & "�п�J�Gchrome, firefox, internet explorer, safari, opera" & vbNewLine & "(���^��p�g)", 0 + 16, "Error")
            Sheets("Web_Infor").Cells(i, "A").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            Sheets("Web_Infor").Cells(i, "A").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
        
    i = i + 1
    Loop Until Sheets("Web_Infor").Cells(i, "A") = ""
    
    
    
    '�T�{BrowserDriverPath
    j = 2
    Do
        If Sheets("Web_Infor").Cells(j, "B") = "" Then
            y = MsgBox("�ж�J" & Sheets("Web_Infor").Cells(j, "A").Value & "��BrowserDriverPath", 0 + 16, "Error")
            Sheets("Web_Infor").Cells(j, "B").Interior.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            If Dir(CStr(Sheets("Web_Infor").Cells(j, "B"))) = "" Then
                x = MsgBox("�䤣��" & Sheets("Web_Infor").Cells(j, "B"), 0 + 16, "Error")
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
    
    '�ˬdSeleniumServerJarPath
    If Sheets("Web_Infor").Cells(2, "F") = "" Then
    
        x = MsgBox("�ж�JSeleniumServerJarPath�ɸ��|" & vbNewLine & "�Ҧp�GC:\Users\Desktop\�ɦW.jar", 0 + 16, "Error")
        CheckAPPandDevice = False
        Exit Function
    Else
        If Dir(CStr(Sheets("Web_Infor").Cells(2, "F"))) = "" Then
        
            x = MsgBox("�䤣��" & Sheets("Web_Infor").Cells(2, "F"), 0 + 16, "Error")
            Sheets("Web_Infor").Cells(2, "F").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
        Else
            Sheets("Web_Infor").Cells(2, "F").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
    
    End If
    
    '�ˬdJarPath
    If Sheets("Web_Infor").Cells(2, "E") = "" Then
    
        x = MsgBox("�ж�JJarPath�ɸ��|" & vbNewLine & "�Ҧp�GC:\Users\Desktop\�ɦW.jar", 0 + 16, "Error")
        CheckAPPandDevice = False
        Exit Function
    Else
    
        If Dir(CStr(Sheets("Web_Infor").Cells(2, "E"))) = "" Then
        
            x = MsgBox("�䤣��" & Sheets("Web_Infor").Cells(2, "E"), 0 + 16, "Error")
            Sheets("Web_Infor").Cells(2, "E").Font.Color = RGB(255, 0, 0)
            CheckAPPandDevice = False
            Exit Function
            
        Else
            Sheets("Web_Infor").Cells(2, "E").Font.Color = RGB(0, 0, 0)
            CheckAPPandDevice = True
        End If
    End If
    
    
End Function
