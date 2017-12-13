Attribute VB_Name = "執行腳本"

Sub RunScript()

    Dim Jar, LaucnhHub, LaunchNode, Json As String
    Dim maxInstances, state As Integer
    maxInstances = 5
    state = 0
    ActiveWorkbook.Save
    Application.Wait Now() + TimeValue("00:00:02") '暫緩2秒，等待Excel存檔
    
    CheckAPPandDeviceResult = CheckAPPandDevice()
    CheckValueResult = CheckValue()
    CheckCommandResult = CheckCommand()
    
    
    If CheckAPPandDeviceResult = True And CheckValueResult = True And CheckCommandResult = True Then
        
        LaucnhHub = "java -jar " & Sheets("Web_Infor").Cells(2, "G") & " -role hub"
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & LaucnhHub, 1) '啟動Selenium Server
        LaunchNode = "java"
        i = 2
        Do
            Select Case Sheets("Web_Infor").Cells(i, "A")
            
            Case "chrome"
                Json = Json & " -browser " & Chr(34) & "browserName=chrome, maxInstances=" & maxInstances & Chr(34) '設定各Browser資訊;Chr(34)為雙引號
                LaunchNode = LaunchNode & " -Dwebdriver.chrome.driver=" & Sheets("Web_Infor").Cells(i, "B") '指定Driver路徑
            Case "firefox"
                Json = Json & " -browser " & Chr(34) & "browserName=firefox, maxInstances=" & maxInstances & Chr(34)
                LaunchNode = LaunchNode & " -Dwebdriver.gecko.driver=" & Sheets("Web_Infor").Cells(i, "B")
            Case "internet explorer"
                Json = Json & " -browser " & Chr(34) & "browserName=internet explorer, maxInstances=" & maxInstances & Chr(34)
                LaunchNode = LaunchNode & " -Dwebdriver.ie.driver=" & Sheets("Web_Infor").Cells(i, "B")
            End Select
        
        i = i + 1
        Loop Until Sheets("Web_Infor").Cells(i, "A") = ""
        LaunchNode = LaunchNode & " -jar " & Sheets("Web_Infor").Cells(2, "G") & " -role node -hub http://localhost:4444/grid/register" & Json '彙整Selenium node指令
        'LaunchNode = LaunchNode & " -jar " & Sheets("Web_Infor").Cells(2, "G") & " -port 5556" & " -role node -hub http://localhost:4444/grid/register" & Json '彙整Selenium node指令
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & LaunchNode, 1) '啟動Selenium node
        
        Application.Wait Now() + TimeValue("00:00:03") '暫緩3秒，等待Server及Node
        Jar = "java -jar " & Sheets("Web_Infor").Cells(2, "F")
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & Jar, 1) '開始測試
        
    End If

End Sub


