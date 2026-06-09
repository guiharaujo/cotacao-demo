Attribute VB_Name = "Utils"
Option Explicit

' =============================================================================
' UTILS — General helpers: clear quotation, format helpers, UI actions
' =============================================================================

' Called by the CLEAR button — asks for confirmation first
Public Sub BotaoLimpar()
    If MsgBox("Clear the current quotation?", vbQuestion + vbYesNo, "Quotation System") = vbNo Then
        Exit Sub
    End If
    LimparCotacao
End Sub

' Clears the QUOTATION sheet without prompting.
' Called by BotaoLimpar and Workbook_Open (always opens fresh).
Public Sub LimparCotacao()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("QUOTATION")
    If ws Is Nothing Then Exit Sub
    Application.EnableEvents = False

    ' Client data
    ws.Range("C8").ClearContents   ' CRM Opp ID
    ws.Range("C10").ClearContents  ' Company name
    ws.Range("J10").ClearContents  ' CNPJ
    ws.Range("N10").ClearContents  ' ZIP code
    ws.Range("C11").ClearContents  ' Address
    ws.Range("L11").ClearContents  ' State
    ws.Range("N11").ClearContents  ' City

    ' Commercial conditions
    ws.Range("C13").ClearContents  ' Payment term
    ws.Range("I13").Value = "Public"
    ws.Range("L13").Value = "YES"
    ws.Range("C14").ClearContents  ' Freight modal
    ws.Range("G14").ClearContents  ' Freight value
    ws.Range("J14").Value = "CIF"
    ws.Range("N14").Value = 7      ' Validity (days)
    ws.Range("C15").ClearContents  ' Lead time
    ws.Range("E15").ClearContents  ' Client conf.
    ws.Range("H15").ClearContents  ' Rep commission
    ws.Range("L15").Value = "NO"   ' 10-year warranty

    ' Product lines (rows 18–37)
    Dim r As Long
    For r = 18 To 37
        ws.Range("B" & r).ClearContents  ' Code
        ws.Range("F" & r).ClearContents  ' Sale price
        ws.Range("D" & r).ClearContents  ' Qty
        ws.Range("K" & r).ClearContents  ' Extra cost
        ws.Range("L" & r).ClearContents  ' Reason
    Next r

    ' Freight modals
    ws.Range("C41").ClearContents
    ws.Range("G41").ClearContents
    ws.Range("K41").ClearContents

    ' Contact
    ws.Range("C44").ClearContents
    ws.Range("K44").ClearContents
    ws.Range("C45").ClearContents

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Logout button on the QUOTATION sheet
Public Sub BotaoSair()
    If MsgBox("Log out and close the quotation?", vbQuestion + vbYesNo, "Quotation System") = vbNo Then
        Exit Sub
    End If
    LimparCotacao
    Logout
    frmLogin.Show
End Sub

' Formats a phone number as (XX) XXXXX-XXXX
Public Function FormatTelefone(ByVal s As String) As String
    Dim d As String, i As Integer
    For i = 1 To Len(s)
        If Mid(s, i, 1) Like "[0-9]" Then d = d & Mid(s, i, 1)
    Next i
    Select Case Len(d)
        Case 11: FormatTelefone = "(" & Left(d, 2) & ") " & Mid(d, 3, 5) & "-" & Right(d, 4)
        Case 10: FormatTelefone = "(" & Left(d, 2) & ") " & Mid(d, 3, 4) & "-" & Right(d, 4)
        Case Else: FormatTelefone = s
    End Select
End Function

' Sanitizes a string for use as a filename (alphanumeric + dash/underscore only)
Public Function SafeFileName(ByVal s As String) As String
    Dim i As Integer, c As String, out As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c Like "[A-Za-z0-9_-]" Then out = out & c
    Next i
    SafeFileName = out
End Function
