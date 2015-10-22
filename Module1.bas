Attribute VB_Name = "Module1"
Private rowNumber As Integer
Private docFolder As String

Sub GeneratePDF()
    UserForm1.Show
End Sub

Function documentPath() As String
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Open document based on semester and fee type
    With ws
        Dim cellTextG As String
        cellTextG = .Range("G" + CStr(rowNumber)).Text
        
        Dim cellTextH As String
        cellTextH = .Range("H" + CStr(rowNumber)).Text
        
        If InStr(cellTextG, "waived") > 0 And InStr(cellTextH, "Spring") > 0 And InStr(cellTextH, "Academic") = 0 Then
            docFolder = "IEPUN"
            documentPath = getCurrentDir + "\semwaived.docx"
        End If
        If InStr(cellTextG, "waived") > 0 And InStr(cellTextH, "Fall 2015 & Spring 2016") > 0 Then
            docFolder = "IEPUN"
            documentPath = getCurrentDir + "\aywaived.docx"
        End If
        If InStr(cellTextG, "paying") > 0 And InStr(cellTextH, "Spring") > 0 And InStr(cellTextH, "Academic") = 0 Then
            docFolder = "VISN"
            documentPath = getCurrentDir + "\semesterfee.docx"
        End If
        If InStr(cellTextG, "paying") > 0 And InStr(cellTextH, "Fall 2015 & Spring 2016") > 0 Then
            docFolder = "VISN"
            documentPath = getCurrentDir + "\ayfee.docx"
        End If
        If InStr(cellTextG, "paying") > 0 And InStr(cellTextH, "Spring 2016 & Fall 2016") > 0 Then
            docFolder = "VISN"
            documentPath = getCurrentDir + "\ayfee.docx"
        End If
        
    End With
    
End Function

Function openAndParse(path As String)
   Dim ws As Worksheet
   Set ws = ActiveSheet

    With ws
    If IsEmpty(.Range("E" + CStr(rowNumber)).Value) Then
        MsgBox .Range("D" + CStr(rowNumber)).Text + " has no ID Number."
    Else
    
   Dim objWord
   Dim objDoc
   Dim objSelection
   Set objWord = CreateObject("Word.Application")
   Set objDoc = objWord.Documents.Open(path)
   Set objSelection = objWord.selection
   
   Dim fNamePath As String
   Dim lNamePath As String

        'parse first name
        objSelection.Find.Text = "Fname"
        objSelection.Find.Forward = True
        objSelection.Find.MatchWholeWord = True
        Const wdReplaceAll = 2
    
        objSelection.Find.Replacement.Text = StrConv(.Range("D" + CStr(rowNumber)).Text, vbProperCase)
        objSelection.Find.Execute , , , , , , , , , , wdReplaceAll
        fNamePath = StrConv(.Range("D" + CStr(rowNumber)).Text, vbProperCase)
        
        'parse last name
        objSelection.Find.Text = "Lname"
        objSelection.Find.Forward = True
        objSelection.Find.MatchWholeWord = True
    
        objSelection.Find.Replacement.Text = StrConv(.Range("C" + CStr(rowNumber)).Text, vbProperCase)
        objSelection.Find.Execute , , , , , , , , , , wdReplaceAll
        lNamePath = StrConv(.Range("C" + CStr(rowNumber)).Text, vbLowerCase)
        
        Dim BString() As String
        BString = Split(.Range("B" + CStr(rowNumber)), "-")
        
        'Parse university name
        objSelection.Find.Text = "Universityname"
        objSelection.Find.Forward = True
        objSelection.Find.MatchWholeWord = True
    
        objSelection.Find.Replacement.Text = BString(0)
        objSelection.Find.Execute , , , , , , , , , , wdReplaceAll
        
        'Parse country name
        objSelection.Find.Text = "Countryname"
        objSelection.Find.Forward = True
        objSelection.Find.MatchWholeWord = True
        
        'Remove parentheses after country and Replace Korea with Republic of Korea
        Dim CSplit() As String
        CSplit = Split(BString(1), "(")
        If InStr(1, CSplit(0), "Korea") > 0 Then
            CSplit(0) = "Republic Of Korea"
        End If
        objSelection.Find.Replacement.Text = CSplit(0)
        objSelection.Find.Execute , , , , , , , , , , wdReplaceAll
        
        'Parse current date
        objSelection.Find.Text = "Date"
        objSelection.Find.Forward = True
        objSelection.Find.MatchWholeWord = True
    
        objSelection.Find.Replacement.Text = Date
        objSelection.Find.Execute , , , , , , , , , , wdReplaceAll
        
        'Parse SBU ID Number
        objSelection.Find.Text = "SBUIDNUM"
        objSelection.Find.Forward = True
        objSelection.Find.MatchWholeWord = True
        
        objSelection.Find.Replacement.Text = .Range("E" + CStr(rowNumber)).Text
        objSelection.Find.Execute , , , , , , , , , , wdReplaceAll
        
        'Saves document as PDF and closes template without saving
        objDoc.SaveAs2 getCurrentDir + "\AcceptanceLetters\" + docFolder + "\" + lNamePath + fNamePath + ".pdf", 17
        objDoc.Close savechanges:=False
        
        objWord.Quit
        Set objDoc = Nothing
        Set objWord = Nothing
    End If

    End With
    
End Function

Function setRowNumber(rn As Integer)
    rowNumber = rn
End Function

Function getRowNumber() As Integer
    getRowNumber = rowNumber
End Function

Function getCurrentDir() As String
    getCurrentDir = Application.ActiveWorkbook.path
End Function












