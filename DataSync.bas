Attribute VB_Name = "DataSync"
Option Explicit

' =============================================================================
' DATASYNC — Loads product and payment data from Azure SQL into the DATA sheet
' =============================================================================
' Uses ADODB CopyFromRecordset (bulk insert) for performance.
' Called automatically on login and manually via the REFRESH button.
' =============================================================================

Public Sub CarregarDados(Optional silent As Boolean = False)
    Dim conn       As Object
    Dim rs         As Object
    Dim ws         As Worksheet
    Dim nProd      As Long
    Dim nPag       As Long
    Dim etapa      As String
    Dim originalVis As Long

    On Error GoTo Falha
    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    Dim cs As String: cs = ConnStr()
    If cs = "" Then GoTo Falha

    etapa = "Connecting to Azure SQL"
    Application.StatusBar = etapa & "..."
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.Open cs

    Set ws = ThisWorkbook.Sheets("DATA")
    originalVis = ws.Visible
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible

    ' --- Clear previous data ---
    etapa = "Clearing DATA sheet"
    Application.StatusBar = etapa & "..."
    ws.Range("A2:AC100000").ClearContents
    ws.Range("AF2:AG10000").ClearContents

    ' --- Load products ---
    etapa = "Querying COT_PRODUTOS"
    Application.StatusBar = etapa & "..."
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3  ' adUseClient (required for CopyFromRecordset)
    rs.Open "SELECT STATUS,CODE,DESCRIPTION,NCM,FAMILY,WATTAGE,COLOR_TEMP,TYPE,COST," & _
            "IE,WEIGHT,HEIGHT,WIDTH,LENGTH,SHARED_BOX,SEG,ICMS,PIS_COFINS,IPI," & _
            "COMMISSION,LOGISTICS,OPERATIONAL,SALES,ADMIN,QUALITY,BENEFIT,DIFAL," & _
            "CONTRIBUTION_MARGIN FROM " & TBL_PRODUTOS & " WHERE STATUS='ACTIVE' ORDER BY CODE", _
            conn, 3, 1

    nProd = rs.RecordCount
    If nProd > 0 Then ws.Range("A2").CopyFromRecordset rs
    rs.Close

    ' --- Load payment terms ---
    etapa = "Querying COT_PAGAMENTO"
    Application.StatusBar = etapa & "..."
    rs.Open "SELECT DESCRIPTION, RATE FROM " & TBL_PAGAMENTO & " ORDER BY DESCRIPTION", _
            conn, 3, 1

    nPag = rs.RecordCount
    If nPag > 0 Then ws.Range("AF2").CopyFromRecordset rs
    rs.Close
    conn.Close

    ' --- Write metadata ---
    ws.Range("AL2").Value = Now()
    ws.Range("AL3").Value = nProd
    ws.Range("AL4").Value = nPag

    ' --- Refresh dynamic dropdowns (non-fatal if fails) ---
    On Error Resume Next
    AtualizarDropdowns
    On Error GoTo Falha

    ws.Visible = originalVis
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    If Not silent Then
        MsgBox "Data refreshed:" & vbCrLf & _
               "  Products:  " & nProd & vbCrLf & _
               "  Payments:  " & nPag, vbInformation, "Quotation System"
    End If
    Exit Sub

Falha:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Visible = originalVis
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    MsgBox "Failed to load data from Azure SQL." & vbCrLf & vbCrLf & _
           "Step: " & etapa & vbCrLf & _
           "Err #" & Err.Number & " (" & Err.Source & ")" & vbCrLf & _
           Err.Description, vbCritical, "Quotation System"
End Sub

' -----------------------------------------------------------------------------
' AtualizarDropdowns
' Rebuilds in-cell validation lists for product codes and payment terms
' based on the current DATA sheet content.
' -----------------------------------------------------------------------------
Private Sub AtualizarDropdowns()
    Dim wsQ As Worksheet: Set wsQ = ThisWorkbook.Sheets("QUOTATION")
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("DATA")

    Dim wasProt As Boolean
    wasProt = wsQ.ProtectContents
    If wasProt Then wsQ.Unprotect ""

    ' Payment terms dropdown
    Dim lastPag As Long
    lastPag = wsD.Cells(wsD.Rows.Count, 32).End(xlUp).Row
    If lastPag >= 2 Then
        wsQ.Range("C13:E13").Validation.Delete
        wsQ.Range("C13:E13").Validation.Add _
            Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="=DATA!$AF$2:$AF$" & lastPag
        wsQ.Range("C13:E13").Validation.IgnoreBlank = True
        wsQ.Range("C13:E13").Validation.InCellDropdown = True
    End If

    ' Product codes dropdown
    Dim lastProd As Long
    lastProd = wsD.Cells(wsD.Rows.Count, 2).End(xlUp).Row
    If lastProd >= 2 Then
        wsQ.Range("B18:B37").Validation.Delete
        wsQ.Range("B18:B37").Validation.Add _
            Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="=DATA!$B$2:$B$" & lastProd
        wsQ.Range("B18:B37").Validation.IgnoreBlank = True
        wsQ.Range("B18:B37").Validation.InCellDropdown = True
    End If

    If wasProt Then
        wsQ.Protect Password:="", DrawingObjects:=True, Contents:=True, _
                    Scenarios:=True, AllowFormattingCells:=True
        wsQ.EnableSelection = 1
    End If
End Sub

' Called by the REFRESH button on the QUOTATION sheet
Public Sub BotaoAtualizar()
    CarregarDados False
End Sub
