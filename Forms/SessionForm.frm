VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SessionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedSession As String ' Declare a public variable to store the session

Private Sub UserForm_Initialize()
    Dim startYear As Integer
    Dim endYear As Integer
    Dim i As Integer

    ' Define the range of session years
    endYear = year(Date)
    startYear = endYear - 50

    ' Populate the ComboBox with session years
    For i = endYear To startYear Step -1
        ComboBox1.AddItem i & "-" & (i + 1)
    Next i
End Sub

Private Sub CommandButton1_Click()
        ' Check if a session is selected
        If ComboBox1.Value = "" Then
            MsgBox "Please select a session year.", vbExclamation
        Else
            SelectedSession = ComboBox1.Value ' Store the selected session in the public variable
            MsgBox "You selected: " & SelectedSession, vbInformation
            Me.Hide ' Close the UserForm after selection
        End If
End Sub
