VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "工作表1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub BorwserList_Change()
    Sheets("Web_Infor").Cells(2, "A") = BorwserList.Text
    Sheets("Web_Infor").Cells(2, "B") = Sheets("Web_InforData").Cells(BorwserList.ListIndex + 2, "B")
End Sub


Private Sub ListBox1_Click()
    ListBox2.clear
    
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.selected(i) = True Then
            j = 1
            Do
                If Sheets(ListBox1.List(i)).Cells(j, "A") = "CaseName" Then
                
                  ListBox2.AddItem (Sheets(ListBox1.List(i)).Cells(j, "B"))
                
                End If
                
            
            j = j + 1
            Loop Until Sheets(ListBox1.List(i)).Cells(j, "A") = ""
        
        End If
    
    Next
    
    Sheets("Web_Infor").Cells(2, "D").clear
    Sheets("Web_Infor").Cells(2, "E").clear
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.selected(i) = True Then
            
            Sheets("Web_Infor").Cells(2, "D").Value = ListBox1.List(i)
        
        End If
    
    Next
End Sub
Sub loadingTestScript()

    ListBox1.clear
    i = 0
    Do
        If ThisWorkbook.Sheets(i + 1).Visible = True And Right(ThisWorkbook.Sheets(i + 1).Name, 11) = "_TestScript" Then
            ListBox1.AddItem (ThisWorkbook.Sheets(i + 1).Name)
        End If
    
    i = i + 1
    Loop Until i = ThisWorkbook.Sheets.Count
    
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Sheets("Web_Infor").Cells(2, "E").clear
    Count = 0
    For i = 0 To ListBox2.ListCount - 1
        
        If ListBox2.selected(i) = True Then
            Count = Count + 1
            
            If Count = 1 Then
                Sheets("Web_Infor").Cells(2, "E").Value = ListBox2.List(i)
            Else
                Sheets("Web_Infor").Cells(2, "E").Value = Sheets("Web_Infor").Cells(2, "E").Value & "," & ListBox2.List(i)
            End If
    
        End If
    
    Next
End Sub

Sub loadingBrowser()

    Dim DeviceUDIDCurrent_index As Integer
    DeviceUDIDCurrent_index = BorwserList.ListIndex
    BorwserList.clear
    i = 2
    Do

        BorwserList.AddItem (Sheets("Web_InforData").Cells(i, "A"))

    i = i + 1
    Loop Until Sheets("Web_InforData").Cells(i, "A") = ""
    BorwserList.Text = Sheets("Web_InforData").Cells(DeviceUDIDCurrent_index + 2, "A")
    '"請選擇瀏覽器"
    
End Sub

Sub loadingURL()
    Dim DeviceUDIDCurrent_index As Integer
    DeviceUDIDCurrent_index = URLList.ListIndex
    URLList.clear
    i = 2
    Do
        URLList.AddItem (Sheets("Web_inforData").Cells(i, "C"))
        
        i = i + 1
    Loop Until Sheets("Web_InforData").Cells(i, "C") = ""
    URLList.Text = Sheets("Web_InforData").Cells(DeviceUDIDCurrent_index + 2, "C")
    '"請選擇測試URL"
End Sub

Private Sub URLList_Change()
    Sheets("Web_Infor").Cells(2, "C") = URLList.Text
End Sub

Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Call loadingURL
    Call loadingTestScript
    Call loadingBrowser
    Application.ScreenUpdating = True
End Sub

