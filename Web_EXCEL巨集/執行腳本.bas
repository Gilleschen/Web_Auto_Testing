Attribute VB_Name = "����}��"

Sub RunScript()

    Dim Jar, LaucnhHub, LaunchNode, Json As String
    Dim maxInstances, state As Integer
    maxInstances = 5
    state = 0
    ActiveWorkbook.Save
    Application.Wait Now() + TimeValue("00:00:02") '�Ƚw2��A����Excel�s��
    
    CheckAPPandDeviceResult = CheckAPPandDevice()
    CheckValueResult = CheckValue()
    CheckCommandResult = CheckCommand()
    
    
    If CheckAPPandDeviceResult = True And CheckValueResult = True And CheckCommandResult = True Then
        
        LaucnhHub = "java -jar " & Sheets("Web_Infor").Cells(2, "G") & " -role hub"
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & LaucnhHub, 1) '�Ұ�Selenium Server
        LaunchNode = "java"
        i = 2
        Do
            Select Case Sheets("Web_Infor").Cells(i, "A")
            
            Case "chrome"
                Json = Json & " -browser " & Chr(34) & "browserName=chrome, maxInstances=" & maxInstances & Chr(34) '�]�w�UBrowser��T;Chr(34)�����޸�
                LaunchNode = LaunchNode & " -Dwebdriver.chrome.driver=" & Sheets("Web_Infor").Cells(i, "B") '���wDriver���|
            Case "firefox"
                Json = Json & " -browser " & Chr(34) & "browserName=firefox, maxInstances=" & maxInstances & Chr(34)
                LaunchNode = LaunchNode & " -Dwebdriver.gecko.driver=" & Sheets("Web_Infor").Cells(i, "B")
            Case "internet explorer"
                Json = Json & " -browser " & Chr(34) & "browserName=internet explorer, maxInstances=" & maxInstances & Chr(34)
                LaunchNode = LaunchNode & " -Dwebdriver.ie.driver=" & Sheets("Web_Infor").Cells(i, "B")
            End Select
        
        i = i + 1
        Loop Until Sheets("Web_Infor").Cells(i, "A") = ""
        LaunchNode = LaunchNode & " -jar " & Sheets("Web_Infor").Cells(2, "G") & " -role node -hub http://localhost:4444/grid/register" & Json '�J��Selenium node���O
        'LaunchNode = LaunchNode & " -jar " & Sheets("Web_Infor").Cells(2, "G") & " -port 5556" & " -role node -hub http://localhost:4444/grid/register" & Json '�J��Selenium node���O
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & LaunchNode, 1) '�Ұ�Selenium node
        
        Application.Wait Now() + TimeValue("00:00:03") '�Ƚw3��A����Server��Node
        Jar = "java -jar " & Sheets("Web_Infor").Cells(2, "F")
        r = Shell(Environ("windir") & "\system32\cmd.exe cmd/k" & Jar, 1) '�}�l����
        
    End If

End Sub


