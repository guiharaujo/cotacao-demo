Attribute VB_Name = "Planilha_QUOTATION"
Option Explicit

' =============================================================================
' QUOTATION SHEET — Worksheet_Change event handler
' =============================================================================
' Behavior:
'   - Typing a product code in B18:B37 → auto-fills minimum price into F (sale price)
'   - Clearing a code → clears sale price, extra cost, and reason
'   - Changing segment (I13), ICMS (L13), commission (H15), or warranty (K15:L15)
'     → recalculates minimum prices for all rows that already have a code
'   - Changing extra cost (K18:K37) → recalculates the affected row
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    If Target.Cells.CountLarge > 50 Then Exit Sub

    Application.EnableEvents = False

    Dim r As Long
    Dim minVal As Variant
    Dim cel As Range

    ' --- Product code entered or cleared (B18:B37) ---
    For Each cel In Target.Cells
        If Not Intersect(cel, Me.Range("B18:B37")) Is Nothing Then
            r = cel.Row
            If Trim(CStr(cel.Value)) = "" Then
                Me.Range("F" & r).ClearContents
                Me.Range("K" & r).ClearContents
                Me.Range("L" & r).ClearContents
            Else
                Application.Calculate
                minVal = Me.Range("E" & r).Value
                If IsNumeric(minVal) Then Me.Range("F" & r).Value = Round(CDbl(minVal), 2)
            End If
        End If
    Next cel

    ' --- Extra cost changed (K18:K37) → recalculate min price for that row ---
    For Each cel In Target.Cells
        If Not Intersect(cel, Me.Range("K18:K37")) Is Nothing Then
            r = cel.Row
            If Trim(CStr(Me.Range("B" & r).Value)) <> "" Then
                Application.Calculate
                minVal = Me.Range("E" & r).Value
                If IsNumeric(minVal) Then Me.Range("F" & r).Value = Round(CDbl(minVal), 2)
            End If
        End If
    Next cel

    ' --- Segment, ICMS, commission, or warranty changed → recalculate all rows ---
    If Not Intersect(Target, Me.Range("I13:J13")) Is Nothing _
       Or Not Intersect(Target, Me.Range("L13")) Is Nothing _
       Or Not Intersect(Target, Me.Range("H15")) Is Nothing _
       Or Not Intersect(Target, Me.Range("K15:L15")) Is Nothing Then
        Application.Calculate
        For r = 18 To 37
            If Trim(CStr(Me.Range("B" & r).Value)) <> "" Then
                minVal = Me.Range("E" & r).Value
                If IsNumeric(minVal) Then Me.Range("F" & r).Value = Round(CDbl(minVal), 2)
            End If
        Next r
    End If

CleanExit:
    Application.EnableEvents = True
End Sub
