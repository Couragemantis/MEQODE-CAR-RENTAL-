VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================
' clsDB.cls
' Centralized Database Connection
'==============================

Option Explicit

' -----------------------------
' Public ADODB connection
' -----------------------------
Public con As ADODB.Connection

' Store record counts per table
Private m_TableCounts As Object   ' Late binding dictionary

Private bIsTyping As Boolean
'==============================
' Open database connection
'==============================
Public Sub OpenDB()
    If con Is Nothing Then
        Set con = New ADODB.Connection
        con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\MEQODES-main\MEQODE\mdb\MasterData.mdb;"
        con.Open
    ElseIf con.State = adStateClosed Then
        con.Open
    End If

    ' Initialize table tracker
    If m_TableCounts Is Nothing Then
        Set m_TableCounts = CreateObject("Scripting.Dictionary")
    End If
End Sub

'==============================
' Close database connection
'==============================
Public Sub CloseDB()
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If

    If Not m_TableCounts Is Nothing Then
        m_TableCounts.RemoveAll
        Set m_TableCounts = Nothing
    End If
End Sub

'==============================
' Safe text for SQL
'==============================
Public Function SafeText(ByVal s As String) As String
    SafeText = Replace(s, "'", "''")
End Function

'==============================
' Detect if table record count changed
'==============================
Public Function HasTableChanged(ByVal tableName As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim currentCount As Long

    If con Is Nothing Then Exit Function
    If con.State <> adStateOpen Then Exit Function

    Set rs = New ADODB.Recordset
    rs.Open "SELECT COUNT(*) AS TotalRec FROM " & tableName, _
            con, adOpenForwardOnly, adLockReadOnly

    currentCount = rs!TotalRec
    rs.Close
    Set rs = Nothing

    ' First time tracking this table
    If Not m_TableCounts.Exists(tableName) Then
        m_TableCounts.Add tableName, currentCount
        HasTableChanged = False
        Exit Function
    End If

    ' Compare with previous count
    If m_TableCounts(tableName) <> currentCount Then
        m_TableCounts(tableName) = currentCount
        HasTableChanged = True
    Else
        HasTableChanged = False
    End If
End Function

