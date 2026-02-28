VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Vehicles"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAutoRefresh 
      Interval        =   2000
      Left            =   12960
      Top             =   5880
   End
   Begin VB.ComboBox comboseat 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3360
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   7440
      TabIndex        =   18
      Top             =   2160
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtprice 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search Plate Number"
      Height          =   855
      Left            =   7440
      TabIndex        =   12
      Top             =   1200
      Width           =   6375
      Begin VB.CommandButton Command9 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Manipulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   2640
      TabIndex        =   7
      Top             =   4920
      Width           =   4455
      Begin VB.CommandButton cmdArchive 
         Caption         =   "&Archive"
         Height          =   495
         Left            =   3360
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Update"
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton ClearInputs 
         Caption         =   "&New"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   4935
   End
   Begin VB.TextBox txtbrand 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox txtplate 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   13680
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   10200
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   10680
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Vehicles"
      BeginProperty Font 
         Name            =   "Segoe UI Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   14535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   668
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Seater"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   668
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   668
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   668
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plate Number"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   195
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   -120
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "frmcar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private db As clsDB
Private rec As ADODB.Recordset
Private isEditing As Boolean
Private bIsTyping As Boolean

Private Sub txtplate_GotFocus(): bIsTyping = True: End Sub
Private Sub txtplate_LostFocus(): bIsTyping = False: End Sub

Private Sub txtbrand_GotFocus(): bIsTyping = True: End Sub
Private Sub txtbrand_LostFocus(): bIsTyping = False: End Sub

Private Sub txtname_GotFocus(): bIsTyping = True: End Sub
Private Sub txtname_LostFocus(): bIsTyping = False: End Sub

Private Sub txtprice_GotFocus(): bIsTyping = True: End Sub
Private Sub txtprice_LostFocus(): bIsTyping = False: End Sub
Public Sub RefreshGrid()
    If Not rec Is Nothing Then
        rec.Requery
        Set DataGrid1.DataSource = rec
        HidecarIDColumn
    End If
End Sub
Private Sub HidecarIDColumn()
    On Error Resume Next
    DataGrid1.Columns(DataGrid1.Columns.Count - 1).Visible = False
End Sub
Private Sub UpdateTextBoxes()
    If rec Is Nothing Then Exit Sub
    If rec.State <> adStateOpen Then Exit Sub
    If rec.BOF And rec.EOF Then Exit Sub
    
    ' Update TextBoxes with current record
    txtplate.Text = rec!Plate
    txtname.Text = rec!Name
    txtbrand.Text = rec!brand
    comboseat.Text = rec!Seater
    txtprice.Text = rec!price
End Sub


Private Sub SetFieldsEditable(Optional ByVal editing As Boolean = False, Optional ByVal adding As Boolean = False)
    If editing And Not adding Then
        txtplate.Locked = True
    Else
        txtplate.Locked = False
    End If
    txtbrand.Locked = False
    txtname.Locked = False
    txtprice.Locked = False
End Sub
Private Sub ClearInputs_Click()
    txtname.Text = ""
    txtbrand.Text = ""
    txtplate.Text = ""
    comboseat.ListIndex = -1
    txtprice.Text = ""
End Sub

Private Sub cmdAdd_Click()
    Dim rsCheck As ADODB.Recordset
    Dim addedBrand As String, addedName As String

    ' =========================
    ' Validate fields
    ' =========================
    If Trim(txtplate.Text) = "" Or Trim(txtbrand.Text) = "" Or Trim(txtname.Text) = "" Or _
       comboseat.ListIndex = -1 Or Trim(txtprice.Text) = "" Then
        MsgBox "Please fill in all required fields!", vbExclamation
        Exit Sub
    End If

    If Not txtplate.Text Like "[A-Za-z][A-Za-z][A-Za-z]####" Then
        MsgBox "Plate Number must be 3 letters followed by 4 numbers!", vbExclamation
        txtplate.SetFocus
        Exit Sub
    End If

    If Len(txtbrand.Text) > 50 Then
        MsgBox "Brand cannot exceed 50 characters!", vbExclamation
        txtbrand.SetFocus
        Exit Sub
    End If

    If Len(txtname.Text) > 50 Then
        MsgBox "Name cannot exceed 50 characters!", vbExclamation
        txtname.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtprice.Text) Or Val(txtprice.Text) <= 0 Then
        MsgBox "Price must be a positive number!", vbExclamation
        txtprice.SetFocus
        Exit Sub
    End If

    ' =========================
    ' Check duplicate Plate
    ' =========================
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT Plate FROM vehicles WHERE Plate = '" & db.SafeText(txtplate.Text) & "'", _
                 db.con, adOpenForwardOnly, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "Plate Number already exists!", vbExclamation
        rsCheck.Close
        Set rsCheck = Nothing
        Exit Sub
    End If

    rsCheck.Close
    Set rsCheck = Nothing

    ' =========================
    ' Add record
    ' =========================
    rec.AddNew
    rec!Plate = txtplate.Text
    rec!brand = txtbrand.Text
    rec!Name = txtname.Text
    rec!Seater = comboseat.Text
    rec!price = Val(txtprice.Text)
    rec.Update

    ' Store brand and name for message
    addedBrand = txtbrand.Text
    addedName = txtname.Text

    ' Refresh grid
    RefreshGrid

    ' Clear form
    txtplate.Text = ""
    txtbrand.Text = ""
    txtname.Text = ""
    comboseat.ListIndex = -1
    txtprice.Text = ""
    txtplate.SetFocus

    ' =========================
    ' Show success message with Brand and Name
    ' =========================
    MsgBox "Vehicle added successfully: " & addedBrand & " - " & addedName, vbInformation
End Sub
Private Sub cmdEdit_Click()
    If rec.EOF Or rec.BOF Then
        MsgBox "Please select a record to edit!", vbExclamation
        Exit Sub
    End If

    ' Validate changes
    If txtbrand.Text = rec!brand And txtname.Text = rec!Name And comboseat.Text = rec!Seater And Val(txtprice.Text) = rec!price Then
        MsgBox "No changes detected!", vbInformation
        Exit Sub
    End If

    ' Confirm edit
    If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Confirm Edit") = vbYes Then
        rec!brand = txtbrand.Text
        rec!Name = txtname.Text
        rec!Seater = comboseat.Text
        rec!price = Val(txtprice.Text)
        rec.Update
        RefreshGrid
    End If
End Sub
Private Sub cmdArchive_Click()
    Dim cnArchive As ADODB.Connection
    Dim rsArchive As ADODB.Recordset

    If rec.EOF Or rec.BOF Then Exit Sub

    ' Archive selected record
    Set cnArchive = New ADODB.Connection
    cnArchive.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\MEQODE\Archives\vehiclearchive\vehiclearchive.mdb;"
    cnArchive.Open

    Set rsArchive = New ADODB.Recordset
    rsArchive.Open "VehicleArchive", cnArchive, adOpenKeyset, adLockOptimistic, adCmdTable

    rsArchive.AddNew
    rsArchive!Plate = rec!Plate
    rsArchive!Name = rec!Name
    rsArchive!brand = rec!brand
    rsArchive!Seater = rec!Seater
    rsArchive!price = rec!price
    rsArchive.Update

    rsArchive.Close
    cnArchive.Close
    Set rsArchive = Nothing
    Set cnArchive = Nothing

    ' Delete from main DB
    rec.Delete
    RefreshGrid ' Refresh after delete

    MsgBox "Record archived successfully!", vbInformation
End Sub
Private Sub Command9_Click()
    Dim searchValue As String
    searchValue = UCase(Trim(Text5.Text))

    If searchValue = "" Then
        rec.Filter = ""
    Else
        rec.Filter = "Plate LIKE '*" & searchValue & "*'"
    End If

    Set DataGrid1.DataSource = rec
    DataGrid1.Refresh
End Sub
Private Sub DataGrid1_Click()
    ' =========================================
    ' Exit if recordset is empty or invalid
    ' =========================================
    If rec Is Nothing Then Exit Sub
    If rec.EOF Or rec.BOF Then Exit Sub

    ' =========================================
    ' Move to the selected row safely
    ' =========================================
    On Error Resume Next
    rec.Bookmark = DataGrid1.Bookmark
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Exit if after Bookmark we are on invalid record
    If rec.EOF Or rec.BOF Then Exit Sub

    ' =========================================
    ' Populate TextBoxes safely (handle NULLs)
    ' =========================================
    txtplate.Text = IIf(IsNull(rec!Plate), "", rec!Plate)
    txtbrand.Text = IIf(IsNull(rec!brand), "", rec!brand)
    txtname.Text = IIf(IsNull(rec!Name), "", rec!Name)
    comboseat.Text = IIf(IsNull(rec!Seater), "", rec!Seater)
   
    
    ' Show numeric value in TextBox (no $ sign in TextBox)
    If IsNull(rec!price) Then
    txtprice.Text = ""
Else
    txtprice.Text = CStr(rec!price)
End If


    ' =========================================
    ' Lock fields if not in editing mode
    ' =========================================
    SetFieldsEditable isEditing
    HidecarIDColumn
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
        KeyCode = 0
    End If
End Sub

    Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
        ' Block typing inside the grid
        KeyAscii = 0
    End Sub
    
Private Sub Form_Load()
    ' Center Form
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    ' Initialize DB
    Set db = New clsDB
    db.OpenDB

    ' Open recordset
    Set rec = New ADODB.Recordset
    With rec
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM vehicles", db.con
    End With

    ' Initialize change tracking
    db.HasTableChanged "vehicles"

    ' Bind to DataGrid
    Set DataGrid1.DataSource = rec
    HidecarIDColumn

    ' Initialize ComboBoxes
    comboseat.Clear
    comboseat.AddItem "4 - Seater"
    comboseat.AddItem "6 - Seater"
    comboseat.AddItem "8 - Seater"

    ' Reset fields
    txtplate.Text = ""
    txtbrand.Text = ""
    txtname.Text = ""
    txtprice.Text = ""
    comboseat.ListIndex = -1

    isEditing = False
    SetFieldsEditable False
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not rec Is Nothing Then
        If rec.State = adStateOpen Then rec.Close
        Set rec = Nothing
    End If

    If Not db Is Nothing Then
        db.CloseDB
        Set db = Nothing
    End If
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
    ' Allow Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Get current cursor position
    Dim pos As Integer
    pos = Text5.SelStart

    ' Limit total length to 7 if not replacing selected text
    If Text5.selLength = 0 And Len(Text5.Text) >= 7 Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If

    ' First 3 characters: letters only, auto-uppercase
    If pos < 3 Then
        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
        ' Convert to uppercase
        KeyAscii = Asc(UCase(Chr(KeyAscii)))

    ' Last 4 characters: numbers only
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    End If
End Sub

Private Sub tmrAutoRefresh_Timer()
    If rec Is Nothing Then Exit Sub
    If rec.State <> adStateOpen Then Exit Sub

    ' Skip refresh while user is typing
    If bIsTyping Then Exit Sub

    If db.HasTableChanged("vehicles") Then
        RefreshGrid
    End If
End Sub
Private Sub txtplate_KeyDown(KeyCode As Integer, Shift As Integer)
' Cancel DELETE key
    If KeyCode = 46 Then
        KeyCode = 0
        Beep
    End If
End Sub

Private Sub txtplate_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    i = Len(txtplate.Text) + 1 ' Position of the character being typed

    ' Always allow Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Limit total length to 7
    If i > 7 Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If

    ' First 3 characters must be letters
    If i <= 3 Then
        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
            Beep
        Else
            ' Convert typed letter to uppercase
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    Else ' Characters 4 to 7 must be numbers
        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)

    ' Allow Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Allow numbers
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub

    ' Allow ONE decimal point only
    If KeyAscii = 46 Then
        If InStr(txtprice.Text, ".") > 0 Then KeyAscii = 0
        Exit Sub
    End If

    ' Block everything else
    KeyAscii = 0

End Sub
