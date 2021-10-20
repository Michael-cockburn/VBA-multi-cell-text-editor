VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Edit_text_form 
   Caption         =   "Edit text"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5505
   OleObjectBlob   =   "Edit_text_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Edit_text_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim box1 As String
    Dim box2 As String
        
    Dim Width1Box As Long
    Dim Height1Box As Long
    Dim Width2Box As Long
    Dim Height2Box As Long
    Dim LeftButton1 As Long
    Dim TopButton1 As Long
    Dim LeftButton2 As Long
    Dim TopButton2 As Long
    Dim HeightRemStartEnd As Long
    
    Dim Frame2Value As String
    Dim Frame3Value As String


Sub userform_initialize()

    FirstBox.Value = ""
    FirstLabel.Visible = False
    FirstBox.Visible = False
    SecondBox.Value = ""
    SecondLabel.Visible = False
    SecondBox.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    
    Edit_text_form.Height = Frame1.Height + 40
    Edit_text_form.Width = Frame1.Width + 25
    
    Width1Box = 180
    Height1Box = 220
    Width2Box = 180
    Height2Box = 270
    HeightRemStartEnd = 400
    
    LeftButton1 = 42
    TopButton1 = 150
    LeftButton2 = 42
    TopButton2 = 180
    
End Sub

Sub AddTextToStart_Click()
    
    FirstBox.Visible = True
    FirstLabel.Visible = True
    SecondLabel.Visible = False
    SecondBox.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    FirstLabel.Caption = "Text to add to start"
    Edit_text_form.Height = Height1Box
    Edit_text_form.Width = Width1Box
    Confirm.Left = LeftButton1
    Confirm.Top = FirstBox.Top + FirstBox.Height + 10
    
End Sub

Private Sub AddTextToEnd_Click()
           
    FirstBox.Visible = True
    FirstLabel.Visible = True
    SecondLabel.Visible = False
    SecondBox.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    FirstLabel.Caption = "Text to add to end"
    Edit_text_form.Height = Height1Box
    Edit_text_form.Width = Width1Box
    Confirm.Left = LeftButton1
    Confirm.Top = FirstBox.Top + FirstBox.Height + 10

End Sub

Private Sub RemoveText_Click()
    
    FirstBox.Visible = True
    FirstLabel.Visible = True
    SecondLabel.Visible = False
    SecondBox.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    FirstLabel.Caption = "Text to remove"
    Edit_text_form.Height = Height1Box
    Edit_text_form.Width = Width1Box
    Confirm.Left = LeftButton1
    Confirm.Top = FirstBox.Top + FirstBox.Height + 10

End Sub

Private Sub ReplaceText_Click()
    
    FirstBox.Visible = True
    FirstLabel.Visible = True
    SecondBox.Visible = True
    SecondLabel.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    FirstLabel.Caption = "text to replace"
    SecondLabel.Caption = "text to replace with"
    Edit_text_form.Height = Height2Box
    Edit_text_form.Width = Width2Box
    Confirm.Left = LeftButton2
    Confirm.Top = SecondBox.Top + SecondBox.Height + 10

End Sub

Private Sub RemovePartOfText_Click()

    FirstBox.Visible = False
    FirstLabel.Visible = False
    SecondBox.Visible = False
    SecondLabel.Visible = False
    Frame2.Visible = True
    Frame3.Visible = True
    Edit_text_form.Width = Frame2.Width + Frame3.Width + 30
    If (Keep = False And Delete = False) Or (Start = False And Last = False) Then
        Confirm.Visible = False
        Frame2.Top = Frame1.Top + Frame1.Height + 10
        Frame3.Top = Frame2.Top
        Edit_text_form.Height = Frame1.Height + Frame2.Height + 50
        Frame4.Visible = False
    ElseIf (Keep = True Or Delete = True) And (Start = True Or Last = True) Then
        Edit_text_form.Height = Frame1.Height + Frame2.Height + Frame4.Height + Confirm.Height + 80
        Confirm.Visible = True
        Frame4.Top = Frame2.Top + Frame2.Height + 10
        Confirm.Top = Frame2.Top + Frame2.Height + Frame4.Height + 20
        Frame4.Visible = True
    End If
    
End Sub

Private Sub Delete_click()
    
    Frame3Value = "Delete"
    DeleteKeep.Caption = Frame3Value & " " & Frame2Value
    NoChars.Left = DeleteKeep.Left + DeleteKeep.Width + 10
    Characters.Left = NoChars.Left + NoChars.Width + 10
    If Start = False And Last = False Then
        Frame4.Visible = False
    Else
        Frame4.Visible = True
        Edit_text_form.Height = Frame1.Height + Frame2.Height + Frame4.Height + Confirm.Height + 80
        Confirm.Visible = True
        Frame4.Top = Frame2.Top + Frame2.Height + 10
        Confirm.Top = Frame2.Top + Frame2.Height + Frame4.Height + 20
    End If

End Sub

Private Sub Keep_click()

    Frame3Value = "Keep"
    Frame4.Visible = True
    DeleteKeep.Caption = Frame3Value & " " & Frame2Value
    NoChars.Left = DeleteKeep.Left + DeleteKeep.Width + 10
    Characters.Left = NoChars.Left + NoChars.Width + 5
    If Start = False And Last = False Then
        Frame4.Visible = False
    Else
        Frame4.Visible = True
        Edit_text_form.Height = Frame1.Height + Frame2.Height + Frame4.Height + Confirm.Height + 80
        Confirm.Visible = True
        Frame4.Top = Frame2.Top + Frame2.Height + 10
        Confirm.Top = Frame2.Top + Frame2.Height + Frame4.Height + 20
    End If
    
End Sub
Private Sub Start_click()

    Frame2Value = "first"
    DeleteKeep.Caption = Frame3Value & " " & Frame2Value
    FirstBox.Visible = False
    FirstLabel.Visible = False
    SecondBox.Visible = False
    SecondLabel.Visible = False
    Frame2.Visible = True
    Frame3.Visible = True
    Confirm.Left = LeftButton2
    Confirm.Top = TopButton2
    If Delete = True Or Keep = True Then Frame4.Visible = True
        NoChars.Left = DeleteKeep.Left + DeleteKeep.Width + 10
        Characters.Left = NoChars.Left + NoChars.Width + 10
    If Delete = False And Keep = False Then
        Frame4.Visible = False
    Else
        Frame4.Visible = True
        Edit_text_form.Height = Frame1.Height + Frame2.Height + Frame4.Height + Confirm.Height + 80
        Confirm.Visible = True
        Frame4.Top = Frame2.Top + Frame2.Height + 10
        Confirm.Top = Frame2.Top + Frame2.Height + Frame4.Height + 20
    End If
    
End Sub


Private Sub Last_click()

    Frame2Value = "last"
    DeleteKeep.Caption = Frame3Value & " " & Frame2Value
    FirstBox.Visible = False
    FirstLabel.Visible = False
    SecondBox.Visible = False
    SecondLabel.Visible = False
    Frame2.Visible = True
    Frame3.Visible = True
    FirstLabel.Caption = "text to replace"
    SecondLabel.Caption = "text to replace with"
    Confirm.Left = LeftButton2
    Confirm.Top = TopButton2
    If Delete = True Or Keep = True Then Frame4.Visible = True
    NoChars.Left = DeleteKeep.Left + DeleteKeep.Width + 10
    Characters.Left = NoChars.Left + NoChars.Width + 10
    If Delete = False And Keep = False Then
        Frame4.Visible = False
    Else
        Frame4.Visible = True
        Edit_text_form.Height = Frame1.Height + Frame2.Height + Frame4.Height + Confirm.Height + 80
        Confirm.Visible = True
        Frame4.Top = Frame2.Top + Frame2.Height + 10
        Confirm.Top = Frame2.Top + Frame2.Height + Frame4.Height + 20
    End If

End Sub


Private Sub NoChars_change()

    NoChars.Width = NoChars.Width + 5
    Characters.Left = NoChars.Left + NoChars.Width + 10
    Frame4.Width = DeleteKeep.Width + NoChars.Width + Characters.Width + 50
    If Frame4.Width + 20 >= Edit_text_form.Width Then
        Edit_text_form.Width = Frame4.Width + 20
    End If
    
End Sub

Private Sub Confirm_Click()

    box1 = FirstBox.Value
    box2 = SecondBox.Value
    
    Dim C As Range
    
    If AddTextToStart = True Then
        For Each C In Selection
        If C.Value <> "" Then C.Value = box1 & C.Value
        Next
    ElseIf AddTextToEnd.Value = True Then
        For Each C In Selection
        If C.Value <> "" Then C.Value = C.Value & box1
        Next
    ElseIf RemoveText.Value = True Then
        For Each C In Selection
        If C.Value <> "" Then C.Value = Replace(C.Value, box1, "")
        Next
    ElseIf ReplaceText.Value = True Then
        For Each C In Selection
        If C.Value <> "" Then C.Value = Replace(C.Value, box1, box2)
        Next
    ElseIf RemovePartOfText = True Then
        
        If NoChars = "" Then
            MsgBox "Please enter a number"
            Exit Sub
        End If
        If Not IsNumeric(NoChars.Value) Then
            MsgBox "Please only enter numbers"
            Exit Sub
        End If
        If Start = True Then
            For Each C In Selection
                If C.Value <> "" And Keep = True Then
                    C.Value = Left(C.Value, NoChars)
                ElseIf C.Value <> "" And Delete = True Then
                    C.Value = Right(C.Value, Len(C.Value) - NoChars)
                End If
            Next
        ElseIf Last = True Then
            For Each C In Selection
                If C.Value <> "" And Delete = True Then
                    C.Value = Left(C.Value, Len(C.Value) - NoChars)
                ElseIf C.Value <> "" And Keep = True Then
                    C.Value = Replace(C.Value, C.Value, Right(C.Value, NoChars))
                End If
            Next
        End If
    End If
        
End Sub
