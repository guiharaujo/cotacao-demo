Attribute VB_Name = "Config"
Option Explicit

' =============================================================================
' CONFIG — Centralized constants and connection settings
' =============================================================================
' IMPORTANT: Never hardcode credentials in production.
' Use environment variables or a secure config file.
' =============================================================================

' Azure SQL table names
Public Const TBL_PRODUTOS   As String = "COT_PRODUTOS"
Public Const TBL_PAGAMENTO  As String = "COT_PAGAMENTO"
Public Const TBL_USUARIOS   As String = "COT_USUARIOS"
Public Const TBL_PROPOSTAS  As String = "COT_PROPOSTAS"
Public Const TBL_CONTADOR   As String = "COT_CONTADOR"

' CRM API (replace with your actual base URL and token)
Public Const CRM_BASE  As String = "https://your-crm.example.com/api/1/"
Public Const CRM_TOKEN As String = "YOUR_CRM_TOKEN_HERE"

' Builds Azure SQL connection string from environment variables.
' Set these variables on your machine before running:
'   COT_DB_SERVER   = your-server.database.windows.net
'   COT_DB_NAME     = your_database
'   COT_DB_USER     = your_user
'   COT_DB_PASSWORD = your_password
Public Function ConnStr() As String
    Dim server   As String: server   = Environ("COT_DB_SERVER")
    Dim database As String: database = Environ("COT_DB_NAME")
    Dim uid      As String: uid      = Environ("COT_DB_USER")
    Dim pwd      As String: pwd      = Environ("COT_DB_PASSWORD")

    If server = "" Or database = "" Or uid = "" Or pwd = "" Then
        MsgBox "Missing environment variables." & vbCrLf & _
               "Set COT_DB_SERVER, COT_DB_NAME, COT_DB_USER, COT_DB_PASSWORD.", _
               vbCritical, "Configuration Error"
        ConnStr = ""
        Exit Function
    End If

    ConnStr = "Provider=SQLOLEDB;" & _
              "Server=" & server & ";" & _
              "Database=" & database & ";" & _
              "UID=" & uid & ";" & _
              "PWD=" & pwd & ";" & _
              "Encrypt=yes;TrustServerCertificate=yes;"
End Function
