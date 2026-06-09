Attribute VB_Name = "Auth"
Option Explicit

' =============================================================================
' AUTH — User authentication and session management
' =============================================================================
' Validates credentials against Azure SQL (COT_USUARIOS table).
' Supports two roles: standard users and management-level approval.
' =============================================================================

Private m_Usuario      As String
Private m_Cargo        As String
Private m_Departamento As String
Private m_Email        As String
Private m_Celular      As String
Private m_UltimoErro   As String

' --- Session getters ---
Public Function UsuarioLogado()  As String: UsuarioLogado  = m_Usuario:      End Function
Public Function CargoLogado()    As String: CargoLogado    = m_Cargo:        End Function
Public Function EmailLogado()    As String: EmailLogado    = m_Email:        End Function
Public Function CelularLogado()  As String: CelularLogado  = m_Celular:      End Function
Public Function UltimoErroLogin() As String: UltimoErroLogin = m_UltimoErro: End Function

' -----------------------------------------------------------------------------
' ValidarLogin
' Validates user/password against COT_USUARIOS.
' Returns True on success and populates session variables.
' -----------------------------------------------------------------------------
Public Function ValidarLogin(ByVal usuario As String, ByVal senha As String) As Boolean
    On Error GoTo Falha
    ValidarLogin = False
    m_UltimoErro = ""

    Dim cs As String: cs = ConnStr()
    If cs = "" Then Exit Function

    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.Open cs

    Dim sql As String
    sql = "SELECT TOP 1 USERNAME, CARGO, DEPARTAMENTO, ATIVO, EMAIL, CELULAR " & _
          "FROM " & TBL_USUARIOS & " " & _
          "WHERE LOWER(USUARIO) = LOWER('" & Replace(usuario, "'", "''") & "') " & _
          "AND SENHA = '" & Replace(senha, "'", "''") & "'"

    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 1

    If Not rs.EOF Then
        If UCase(Trim(CStr(rs.Fields("ATIVO").Value))) = "SIM" Then
            m_Usuario      = CStr(rs.Fields("USERNAME").Value)
            m_Cargo        = CStr(rs.Fields("CARGO").Value)
            m_Departamento = CStr(rs.Fields("DEPARTAMENTO").Value)
            m_Email        = CStr("" & rs.Fields("EMAIL").Value)
            m_Celular      = CStr("" & rs.Fields("CELULAR").Value)
            ValidarLogin   = True
        Else
            m_UltimoErro = "Inactive user."
        End If
    Else
        m_UltimoErro = "Invalid username or password."
    End If

    rs.Close
    conn.Close
    Exit Function
Falha:
    m_UltimoErro = "Error #" & Err.Number & ": " & Err.Description
    ValidarLogin = False
End Function

' -----------------------------------------------------------------------------
' ValidarDiretoria
' Same as ValidarLogin but requires CARGO LIKE '%Management%'.
' Used for approval workflows above the seller's authority level.
' -----------------------------------------------------------------------------
Public Function ValidarDiretoria(ByVal usuario As String, ByVal senha As String) As Boolean
    On Error GoTo Falha
    ValidarDiretoria = False
    m_UltimoErro = ""

    Dim cs As String: cs = ConnStr()
    If cs = "" Then Exit Function

    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.Open cs

    Dim sql As String
    sql = "SELECT TOP 1 USERNAME, CARGO, ATIVO FROM " & TBL_USUARIOS & " " & _
          "WHERE LOWER(USUARIO) = LOWER('" & Replace(usuario, "'", "''") & "') " & _
          "AND SENHA = '" & Replace(senha, "'", "''") & "' " & _
          "AND UPPER(CARGO) LIKE '%MANAGEMENT%' AND ATIVO='SIM'"

    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 1

    If Not rs.EOF Then
        ValidarDiretoria = True
    Else
        m_UltimoErro = "User not authorized (requires Management role)."
    End If

    rs.Close
    conn.Close
    Exit Function
Falha:
    m_UltimoErro = "Error #" & Err.Number & ": " & Err.Description
    ValidarDiretoria = False
End Function

' -----------------------------------------------------------------------------
' PosLogin — runs after successful login
' Reveals working sheets, writes session data, loads products from Azure SQL.
' -----------------------------------------------------------------------------
Public Sub PosLogin()
    Application.ScreenUpdating = False
    Dim wb As Workbook: Set wb = ThisWorkbook

    wb.Sheets("QUOTATION").Visible = xlSheetVisible
    wb.Sheets("FREIGHT").Visible   = xlSheetVisible
    wb.Sheets("DATA").Visible      = xlSheetVeryHidden

    With wb.Sheets("DATA")
        .Range("AI2").Value = m_Usuario
        .Range("AJ2").Value = m_Cargo
        .Range("AI3").Value = m_Email
        .Range("AI4").Value = m_Celular
    End With

    CarregarDados True

    wb.Sheets("LOGIN").Visible = xlSheetVeryHidden
    wb.Sheets("QUOTATION").Activate
    Application.ScreenUpdating = True
End Sub

' -----------------------------------------------------------------------------
' Logout — clears session and returns to login screen
' -----------------------------------------------------------------------------
Public Sub Logout()
    Application.ScreenUpdating = False
    m_Usuario = "": m_Cargo = "": m_Departamento = ""
    m_Email   = "": m_Celular = ""

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = IIf(ws.Name = "LOGIN", xlSheetVisible, xlSheetVeryHidden)
    Next ws

    ThisWorkbook.Sheets("LOGIN").Activate
    Application.ScreenUpdating = True
End Sub
