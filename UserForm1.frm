VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Study Abroad PDF Generation"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    Dim inputTxt As String
    
    inputTxt = TextBox1.Text
    'Check for non-numeric input
    If Not IsNumeric(inputTxt) And InStr(inputTxt, "to") = 0 Then
        MsgBox "Not a number"
    Else
    If InStr(inputTxt, "to") = 0 Then
        setRowNumber (CInt(inputTxt))
        If getRowNumber < 2 Or getRowNumber > ActiveSheet.UsedRange.Rows.Count Then
            MsgBox "Invalid row number"
        Else
            openAndParse (documentPath())
        End If
        
    Else
        'Check for if keyword to generate pdf's for multiple rows
         Dim spl() As String
         spl = Split(inputTxt, "to")
         Dim numLeft As Integer
         Dim numRight As Integer
         
         numLeft = CInt(spl(0))
         numRight = CInt(spl(1))
         If numLeft < 2 Or numLeft > numRight Or numLeft > ActiveSheet.UsedRange.Rows.Count Or numRight > ActiveSheet.UsedRange.Rows.Count Then
            MsgBox "Invalid Range"
         Else
            Dim i As Integer
            For i = numLeft To numRight
                setRowNumber (i)
                openAndParse (documentPath())
            Next i
         End If
    End If
    
    End If
End Sub

