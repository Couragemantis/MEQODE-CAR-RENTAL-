VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBooking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Booking"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtReturn 
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   132579329
      CurrentDate     =   46078
   End
   Begin MSComCtl2.DTPicker dtPick 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   132579329
      CurrentDate     =   46078
   End
   Begin VB.TextBox txtSearchBrand 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdBook 
      Caption         =   "&Book Now"
      Height          =   615
      Left            =   840
      TabIndex        =   14
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtDays 
      Height          =   375
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox cboPrice 
      Height          =   315
      Left            =   12600
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox cboSeater 
      Height          =   315
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox cboCarName 
      Height          =   315
      Left            =   4560
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox cboBrand 
      Height          =   315
      Left            =   4560
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox cboPlate 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtSearchCarName 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox cboCustomerName 
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cboLicense 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtSearchCustomer 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Return Date"
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Pick Up Date"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lbllbrand 
      Caption         =   "Vehicle Brand"
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblVehicle 
      Caption         =   "Vehicle Name"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblCustomer 
      Caption         =   "Customer"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As clsDB

Private Function SafeText(ByVal s As String) As String
    SafeText = Replace(s, "'", "''")
End Function

Private Sub dtPick_Validate(Cancel As Boolean)
 If dtPick.Value < Date Then
        MsgBox "Pickup date cannot be in the past.", vbExclamation, "Invalid Date"
        dtPick.Value = Date
        Cancel = True
    End If

End Sub
Private Sub dtReturn_Validate(Cancel As Boolean)
If dtReturn.Value < Date Then
        MsgBox "Return date cannot be in the past.", vbExclamation, "Invalid Date"
        dtReturn.Value = Date
        Cancel = True
    End If
End Sub

'=========================
' Form Load
'=========================
Private Sub Form_Load()
        Randomize       ' Initialize random number generator
    Set db = New clsDB
    db.OpenDB
UpdateBookingStatus
    LoadCustomerList
    LoadVehicleList
    UpdateComputation
End Sub


Private Sub txtSearchCarName_Change()
    FilterVehicles txtSearchCarName.Text, txtSearchBrand.Text
End Sub

Private Sub txtSearchBrand_Change()
    FilterVehicles txtSearchCarName.Text, txtSearchBrand.Text
End Sub

Sub FilterVehicles(Optional ByVal carName As String = "", Optional ByVal brand As String = "")
    Dim rs As ADODB.Recordset
    Dim sql As String

    Set rs = New ADODB.Recordset

    ' Only filter by Name and Brand
    sql = "SELECT * FROM vehicles WHERE Name LIKE '%" & carName & "%' AND Brand LIKE '%" & brand & "%'"

    rs.CursorLocation = adUseClient
    rs.Open sql, db.con, adOpenStatic, adLockReadOnly

    ' Clear combo boxes
    cboPlate.Clear
    cboBrand.Clear
    cboCarName.Clear
    cboSeater.Clear
    cboPrice.Clear

    ' Add results
    Do While Not rs.EOF
        AddVehicleRow rs
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub
Sub LoadCustomerList(Optional ByVal keyword As String = "")
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim sLicense As String, sName As String
    Dim lCusID As Long

    Set rs = New ADODB.Recordset

    ' Only search by Name
    sql = "SELECT License, Name, cusID FROM customer " & _
          "WHERE Name LIKE '%" & keyword & "%' ORDER BY Name"

    rs.CursorLocation = adUseClient
    rs.Open sql, db.con, adOpenStatic, adLockReadOnly

    cboLicense.Clear
    cboCustomerName.Clear

    Do While Not rs.EOF
        ' Null-safe conversion
        If IsNull(rs.Fields("License").Value) Then
            sLicense = ""
        Else
            sLicense = CStr(rs.Fields("License").Value)
        End If

        If IsNull(rs.Fields("Name").Value) Then
            sName = ""
        Else
            sName = CStr(rs.Fields("Name").Value)
        End If

        If IsNull(rs.Fields("cusID").Value) Then
            lCusID = 0
        Else
            lCusID = CLng(rs.Fields("cusID").Value)
        End If

        cboLicense.AddItem sLicense
        cboLicense.ItemData(cboLicense.NewIndex) = lCusID

        cboCustomerName.AddItem sName
        cboCustomerName.ItemData(cboCustomerName.NewIndex) = lCusID

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub txtSearchCustomer_Change()
    Dim keyword As String
    keyword = txtSearchCustomer.Text
    LoadCustomerList keyword
End Sub
Sub SyncCustomer(ByVal id As Long)
    Dim i As Integer

    For i = 0 To cboLicense.ListCount - 1
        If cboLicense.ItemData(i) = id Then cboLicense.ListIndex = i
    Next

    For i = 0 To cboCustomerName.ListCount - 1
        If cboCustomerName.ItemData(i) = id Then cboCustomerName.ListIndex = i
    Next
End Sub

Private Sub cboLicense_Click()
    If cboLicense.ListIndex < 0 Then Exit Sub
    SyncCustomer cboLicense.ItemData(cboLicense.ListIndex)
End Sub

Private Sub cboCustomerName_Click()
    If cboCustomerName.ListIndex < 0 Then Exit Sub
    SyncCustomer cboCustomerName.ItemData(cboCustomerName.ListIndex)
End Sub

'=========================
' Load Vehicles
'=========================
Sub LoadVehicleList(Optional ByVal keyword As String = "")
    Dim rs As ADODB.Recordset
    Dim sql As String

    Set rs = New ADODB.Recordset

    sql = "SELECT * FROM vehicles WHERE Plate LIKE '%" & keyword & "%' OR Brand LIKE '%" & keyword & "%' OR Name LIKE '%" & keyword & "%' OR Seater LIKE '%" & keyword & "%'"

    rs.CursorLocation = adUseClient
    rs.Open sql, db.con, adOpenStatic, adLockReadOnly

    cboPlate.Clear
    cboBrand.Clear
    cboCarName.Clear
    cboSeater.Clear
    cboPrice.Clear

    Do While Not rs.EOF
        AddVehicleRow rs
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Sub AddVehicleRow(rs As ADODB.Recordset)
    Dim id As Long
    Dim sPlate As String, sBrand As String, sName As String, sSeater As String
    Dim dPrice As Double
    Static colPlate As Collection, colBrand As Collection, colCarName As Collection
    Static colSeater As Collection, colPrice As Collection

    ' Initialize collections once
    If colPlate Is Nothing Then Set colPlate = New Collection
    If colBrand Is Nothing Then Set colBrand = New Collection
    If colCarName Is Nothing Then Set colCarName = New Collection
    If colSeater Is Nothing Then Set colSeater = New Collection
    If colPrice Is Nothing Then Set colPrice = New Collection

    ' Null-safe conversion
    sPlate = IIf(IsNull(rs.Fields("Plate").Value), "", CStr(rs.Fields("Plate").Value))
    sBrand = IIf(IsNull(rs.Fields("Brand").Value), "", CStr(rs.Fields("Brand").Value))
    sName = IIf(IsNull(rs.Fields("Name").Value), "", CStr(rs.Fields("Name").Value))
    sSeater = IIf(IsNull(rs.Fields("Seater").Value), "", CStr(rs.Fields("Seater").Value))
    dPrice = IIf(IsNull(rs.Fields("Price").Value), 0, CDbl(rs.Fields("Price").Value))
    id = IIf(IsNull(rs.Fields("carID").Value), 0, CLng(rs.Fields("carID").Value))

    '========================
    ' Add unique Plate
    On Error Resume Next
    colPlate.Add sPlate, sPlate
    If Err.Number = 0 Then
        cboPlate.AddItem sPlate
        cboPlate.ItemData(cboPlate.NewIndex) = id
    End If
    Err.Clear

    ' Add unique Brand
    colBrand.Add sBrand, sBrand
    If Err.Number = 0 Then
        cboBrand.AddItem sBrand
        cboBrand.ItemData(cboBrand.NewIndex) = id
    End If
    Err.Clear

    ' Add unique Car Name
    colCarName.Add sName, sName
    If Err.Number = 0 Then
        cboCarName.AddItem sName
        cboCarName.ItemData(cboCarName.NewIndex) = id
    End If
    Err.Clear

    ' Add unique Seater
    colSeater.Add sSeater, sSeater
    If Err.Number = 0 Then
        cboSeater.AddItem sSeater
        cboSeater.ItemData(cboSeater.NewIndex) = id
    End If
    Err.Clear

    ' Add unique Price
    colPrice.Add dPrice, CStr(dPrice)
    If Err.Number = 0 Then
        cboPrice.AddItem Format(dPrice, "0.00")
        cboPrice.ItemData(cboPrice.NewIndex) = id
    End If
    Err.Clear

    On Error GoTo 0
End Sub

Sub UpdateBookingStatus()
    Dim sql As String
    sql = "UPDATE bookings " & _
          "SET Status = IIf(Date() < PickDate, 'Reserved', IIf(Date() <= ReturnDate, 'On Going', 'Overdue'))"
    db.con.Execute sql
End Sub
Sub SyncVehicle(ByVal id As Long)
    Dim cboList(4) As ComboBox
    Dim i As Integer, j As Integer

    ' Assign the combo boxes to the typed array
    Set cboList(0) = cboPlate
    Set cboList(1) = cboBrand
    Set cboList(2) = cboCarName
    Set cboList(3) = cboSeater
    Set cboList(4) = cboPrice

    ' Loop through each combo box
    For i = 0 To 4
        For j = 0 To cboList(i).ListCount - 1
            If cboList(i).ItemData(j) = id Then
                cboList(i).ListIndex = j
                Exit For  ' Stop after first match
            End If
        Next j
    Next i

    ' Update computation (price x days)
    UpdateComputation
End Sub
Private Sub cboPlate_Click()
    If cboPlate.ListIndex >= 0 Then
        SyncVehicle cboPlate.ItemData(cboPlate.ListIndex)
        UpdateComputation
    End If
End Sub
Private Sub cboBrand_Click()
    If cboBrand.ListIndex >= 0 Then
        SyncVehicle cboBrand.ItemData(cboBrand.ListIndex)
        UpdateComputation
    End If
End Sub
Private Sub cboCarName_Click()
    If cboCarName.ListIndex >= 0 Then
        SyncVehicle cboCarName.ItemData(cboCarName.ListIndex)
        UpdateComputation
    End If
End Sub
Private Sub cboSeater_Click()
    If cboSeater.ListIndex >= 0 Then
        SyncVehicle cboSeater.ItemData(cboSeater.ListIndex)
        UpdateComputation
    End If
End Sub
Private Sub cboPrice_Click()
    If cboPrice.ListIndex >= 0 Then
        SyncVehicle cboPrice.ItemData(cboPrice.ListIndex)
    End If
End Sub
Sub UpdateComputation()

    On Error Resume Next

    Dim pickD As Date, returnD As Date, today As Date
    Dim plannedDays As Long, actualDays As Long
    Dim status As String
    Dim carID As Long, price As Double
    Dim totalDays As Long

    today = Date

    ' Validate dates
    If IsDate(dtPick.Value) = False Or IsDate(dtReturn.Value) = False Then Exit Sub

    pickD = dtPick.Value
    returnD = dtReturn.Value

    ' Planned booking days (minimum 1)
    plannedDays = DateDiff("d", pickD, returnD)
    If plannedDays <= 0 Then plannedDays = 1

    ' ===== STATUS LOGIC =====
    If today < pickD Then
        status = "Reserved"
        totalDays = plannedDays

    ElseIf today >= pickD And today <= returnD Then
        status = "On Going"
        actualDays = DateDiff("d", pickD, today)
        If actualDays <= 0 Then actualDays = 1
        totalDays = actualDays

    Else
        status = "Overdue"
        actualDays = DateDiff("d", pickD, today)
        If actualDays <= 0 Then actualDays = 1
        totalDays = actualDays
    End If

    ' ===== GET CAR PRICE =====
    carID = 0
    If cboPlate.ListIndex >= 0 Then
        carID = cboPlate.ItemData(cboPlate.ListIndex)
    End If

    price = GetCarPrice(carID)

    ' ===== OUTPUT =====
    txtDays.Text = totalDays
    txtStatus.Text = status
    txtTotal.Text = Format(totalDays * price, "0.00")

End Sub
' Function to compute dynamic status
Function ComputeStatus(pickD As Date, returnD As Date, Optional today As Date) As String
    If today = 0 Then today = Date
    
    If today < pickD Then
        ComputeStatus = "Reserved"
    ElseIf today >= pickD And today <= returnD Then
        ComputeStatus = "On Going"
    Else
        ComputeStatus = "Overdue"
    End If
End Function
Private Sub dtPick_Change()
    UpdateComputation
End Sub
' Get customer contact from cusID
Function GetCustomerContact(cusID As Long) As String
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT Contact FROM customer WHERE cusID=" & cusID, db.con, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then GetCustomerContact = rs.Fields("Contact").Value Else GetCustomerContact = ""
    rs.Close
    Set rs = Nothing
End Function

' Get customer type from cusID
Function GetCustomerType(cusID As Long) As String
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT Type FROM customer WHERE cusID=" & cusID, db.con, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then GetCustomerType = rs.Fields("Type").Value Else GetCustomerType = ""
    rs.Close
    Set rs = Nothing
End Function

' Get customer expiration from cusID
Function GetCustomerExpiration(cusID As Long) As Date
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT Expiration FROM customer WHERE cusID=" & cusID, db.con, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then GetCustomerExpiration = rs.Fields("Expiration").Value Else GetCustomerExpiration = Date
    rs.Close
    Set rs = Nothing
End Function

' Get vehicle price from carID
Function GetCarPrice(carID As Long) As Double
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT Price FROM vehicles WHERE carID=" & carID, db.con, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then GetCarPrice = rs.Fields("Price").Value Else GetCarPrice = 0
    rs.Close
    Set rs = Nothing
End Function
Private Sub dtReturn_Change()
    UpdateComputation
End Sub
Private Sub cmdBook_Click()

    Dim code As String
    Dim days As Long
    Dim total As Double
    Dim cusID As Long, carID As Long
    Dim cusContact As String, cusType As String
    Dim cusExp As Date
    Dim carPricePerDay As Double
    Dim sql As String

    '-----------------------------
    ' Validate selection
    '-----------------------------
    If cboLicense.ListIndex < 0 Or cboPlate.ListIndex < 0 Then
        MsgBox "Select a customer and vehicle first.", vbExclamation
        Exit Sub
    End If

    '-----------------------------
    ' IDs
    '-----------------------------
    cusID = cboLicense.ItemData(cboLicense.ListIndex)
    carID = cboPlate.ItemData(cboPlate.ListIndex)

    '-----------------------------
    ' Generate booking code
    '-----------------------------
    code = GenerateBookingCode()

    '-----------------------------
    ' Computed values
    '-----------------------------
    days = CLng(txtDays.Text)
    total = CDbl(txtTotal.Text)

    '-----------------------------
    ' Get customer info
    '-----------------------------
    cusContact = GetCustomerContact(cusID)
    cusType = GetCustomerType(cusID)
    cusExp = GetCustomerExpiration(cusID)

    '-----------------------------
    ' Get vehicle info
    '-----------------------------
    carPricePerDay = GetCarPrice(carID)

    '-----------------------------
    ' BUILD SQL  (16 fields ONLY)
    '-----------------------------
    sql = "INSERT INTO bookings (" & _
          "bookingCode, CusLicense, CusName, CusContact, CusType, CusExpiration, " & _
          "CarPlate, CarBrand, CarName, CarSeater, CarPrice, " & _
          "PickDate, ReturnDate, Days, Status, TotalPrice) VALUES (" & _
          "'" & SafeText(code) & "'," & _
          "'" & SafeText(cboLicense.Text) & "'," & _
          "'" & SafeText(cboCustomerName.Text) & "'," & _
          "'" & SafeText(cusContact) & "'," & _
          "'" & SafeText(cusType) & "'," & _
          "#" & Format(cusExp, "mm/dd/yyyy") & "#," & _
          "'" & SafeText(cboPlate.Text) & "'," & _
          "'" & SafeText(cboBrand.Text) & "'," & _
          "'" & SafeText(cboCarName.Text) & "'," & _
          "'" & SafeText(cboSeater.Text) & "'," & _
          carPricePerDay & "," & _
          "#" & Format(dtPick.Value, "mm/dd/yyyy") & "#," & _
          "#" & Format(dtReturn.Value, "mm/dd/yyyy") & "#," & _
          days & "," & _
          "'" & SafeText(txtStatus.Text) & "'," & _
          total & ")"

    '-----------------------------
    ' Execute
    '-----------------------------
    On Error GoTo ErrHandler
    db.con.Execute sql

    MsgBox "Booking saved! Code: " & code, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error saving booking: " & Err.Description & vbCrLf & vbCrLf & sql, vbCritical

End Sub
Function GenerateBookingCode() As String
    Dim code As String, rs As New ADODB.Recordset

Retry:
    code = "BOOK-" & Int((9999 - 1000 + 1) * Rnd + 1000) & Chr(Int((90 - 65 + 1) * Rnd + 65))
    rs.CursorLocation = adUseClient
    rs.Open "SELECT bookingCode FROM bookings WHERE bookingCode='" & code & "'", db.con, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        GoTo Retry
    End If
    rs.Close
    Set rs = Nothing

    GenerateBookingCode = code
End Function

