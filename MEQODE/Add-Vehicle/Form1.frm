VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Manage Vehicles"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comboseat 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3360
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   7440
      TabIndex        =   23
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
      TabIndex        =   20
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search Plate Number"
      Height          =   855
      Left            =   7440
      TabIndex        =   17
      Top             =   1200
      Width           =   6375
      Begin VB.CommandButton Command9 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4800
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Navigator"
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
      Left            =   7440
      TabIndex        =   8
      Top             =   4920
      Width           =   3975
      Begin VB.CommandButton Command8 
         Caption         =   "&Last"
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Next"
         Height          =   495
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Previous"
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&First"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
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
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Update"
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton ClearInputs 
         Caption         =   "&New"
         Height          =   495
         Left            =   120
         TabIndex        =   9
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
      TabIndex        =   22
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
      TabIndex        =   21
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim isEditing As Boolean


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
    txtbrand.Text = rec!Brand
    cmbseater.Text = rec!Seater
    txtprice.Text = rec!Price
End Sub


Private Sub SetFieldsEditable(Optional ByVal editing As Boolean = False, Optional ByVal adding As Boolean = False)
    ' Text1 (Student ID) behavior:
    ' - Locked when editing an existing record
    ' - Editable when adding a new record
    If editing And Not adding Then
        txtplate.Locked = True
    Else
        txtplate.Locked = False
    End If

    ' Other fields
    txtplate.Locked = False
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
    
    ' Validate required fields
    If Trim(txtplate.Text) = "" Or _
       Trim(txtbrand.Text) = "" Or _
       Trim(txtname.Text) = "" Or _
       comboseat.ListIndex = -1 Or _
       Trim(txtprice.Text) = "" Then
       
        MsgBox "Please fill in all required fields!", vbExclamation
        Exit Sub
    End If

    ' Validate Plate (ABC1234)
    If Not txtplate.Text Like "[A-Za-z][A-Za-z][A-Za-z]####" Then
        MsgBox "Plate Number must be 3 letters followed by 4 numbers!", vbExclamation
        txtplate.SetFocus
        Exit Sub
    End If

    ' Validate Text Length
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

    ' Validate Price
    
    ' Check duplicate Plate
    Dim rsCheck As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT Plate FROM vehicles WHERE Plate = '" & Replace(txtplate.Text, "'", "''") & "'", _
                 con, adOpenForwardOnly, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "Plate Number already exists!", vbExclamation
        rsCheck.Close
        Set rsCheck = Nothing
        Exit Sub
    End If

    rsCheck.Close
    Set rsCheck = Nothing

    ' Add new record
    rec.AddNew
    rec!Plate = txtplate.Text
    rec!Brand = txtbrand.Text
    rec!Name = txtname.Text
    rec!Seater = comboseat.Text
    rec!Price = txtprice.Text
        
    rec.Update

    ' Refresh grid and format Price column
    rec.Requery
    Set DataGrid1.DataSource = rec
    HidecarIDColumn
   

    ' Clear form
    txtplate.Text = ""
    txtbrand.Text = ""
    txtname.Text = ""
    comboseat.ListIndex = -1
    txtprice.Text = ""
    txtplate.SetFocus

    isEditing = False
    SetFieldsEditable False

    MsgBox "Record added successfully!", vbInformation

End Sub

Private Sub cmdEdit_Click()

    If rec.EOF Or rec.BOF Then
        MsgBox "Please select a record to edit!", vbExclamation
        Exit Sub
    End If

    ' Validate required fields
    If Trim(txtplate.Text) = "" Or _
       Trim(txtbrand.Text) = "" Or _
       Trim(txtname.Text) = "" Or _
       comboseat.ListIndex = -1 Or _
       Trim(txtprice.Text) = "" Then
       
        MsgBox "Please fill in all required fields!", vbExclamation
        Exit Sub
    End If

    ' Validate Plate (ABC1234)
    If Not txtplate.Text Like "[A-Za-z][A-Za-z][A-Za-z]####" Then
        MsgBox "Plate Number must be 3 letters followed by 4 numbers!", vbExclamation
        txtplate.SetFocus
        Exit Sub
    End If

    ' Validate Text Length
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

    ' Validate Price
   
    ' Check if anything changed
    If txtbrand.Text = rec!Brand And _
       txtname.Text = rec!Name And _
       comboseat.Text = rec!Seater And _
       (txtprice.Text) = rec!Price Then
       
        MsgBox "No changes detected!", vbInformation
        Exit Sub
    End If

    ' Confirm save
    If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Confirm Edit") = vbYes Then

        rec!Plate = txtplate.Text
        rec!Brand = txtbrand.Text
        rec!Name = txtname.Text
        rec!Seater = comboseat.Text
        rec!Price = (txtprice.Text)   ' numeric only
        rec.Update

        ' Refresh grid and format Price column
        rec.Requery
        Set DataGrid1.DataSource = rec
        HidecarIDColumn

        MsgBox "Record updated successfully!", vbInformation

        ' Clear form
        txtplate.Text = ""
        txtbrand.Text = ""
        txtname.Text = ""
        comboseat.ListIndex = -1
        txtprice.Text = ""

        isEditing = False
        SetFieldsEditable False
    End If

End Sub

Private Sub cmdArchive_Click()
    Dim cnArchive As ADODB.Connection
    Dim rsArchive As ADODB.Recordset
    Dim selectedID As Long
    
    ' ===============================
    ' Ensure a row is selected
    ' ===============================
    If rec Is Nothing Then Exit Sub
    If rec.EOF Or rec.BOF Then
        MsgBox "Please select a record first!", vbExclamation
        Exit Sub
    End If

    ' ===============================
    ' Get carID of the selected row
    ' ===============================
    selectedID = rec!carID

    ' ===============================
    ' Open connection to archive MDB
    ' ===============================
    Set cnArchive = New ADODB.Connection
    cnArchive.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\MEQODE\Archives\vehiclearchive\vehiclearchive.mdb;"

    ' ===============================
    ' Open archive table recordset
    ' ===============================
    Set rsArchive = New ADODB.Recordset
    rsArchive.Open "VehicleArchive", cnArchive, adOpenKeyset, adLockOptimistic, adCmdTable

    ' ===============================
    ' Copy selected row to archive
    ' ===============================
    rsArchive.AddNew
    rsArchive!Plate = rec!Plate
    rsArchive!Name = rec!Name
    rsArchive!Brand = rec!Brand
    rsArchive!Seater = rec!Seater
    rsArchive!Price = rec!Price
    rsArchive.Update

    ' ===============================
    ' Delete from main table (optional)
    ' ===============================
    rec.Delete
    rec.Update  ' If using batch update, use rec.UpdateBatch

    ' ===============================
    ' Refresh datagrid
    ' ===============================
    rec.Requery

    ' ===============================
    ' Close connections
    ' ===============================
    rsArchive.Close
    cnArchive.Close
    Set rsArchive = Nothing
    Set cnArchive = Nothing

    MsgBox "Record archived successfully!", vbInformation
End Sub

Private Sub Command5_Click()

If txtplate.Text = "" Then
Exit Sub
End If

If txtbrand.Text = "" Then
Exit Sub
End If

If txtname.Text = "" Then
Exit Sub
End If

If cmbseater.Text = "" Then
Exit Sub
End If

If txtprice.Text = "" Then
Exit Sub
End If

 If rec Is Nothing Then
 Exit Sub
 End If
    If rec.State <> adStateOpen Then Exit Sub
    If rec.BOF And rec.EOF Then
    Exit Sub
    End If
    rec.MoveFirst
    UpdateTextBoxes
    
    Set DataGrid1.DataSource = rec
HidecarIDColumn
End Sub

Private Sub Command6_Click()

    ' Validate that required fields are not empty
    If txtplate.Text = "" Or txtbrand.Text = "" Or txtname.Text = "" Or _
       cmbseater.Text = "" Or txtprice.Text = "" Then Exit Sub

    ' Ensure recordset exists and is open
    If rec Is Nothing Then Exit Sub
    If rec.State <> adStateOpen Then Exit Sub
    If rec.BOF And rec.EOF Then Exit Sub   ' no records

    ' Move to previous record
    If Not rec.BOF Then
        rec.MovePrevious

        ' Make sure we don’t go before first record
        If rec.BOF Then
            rec.MoveFirst
            MsgBox "Already at the first record.", vbInformation
        End If

        ' Update textboxes with current record
        UpdateTextBoxes
    Else
        MsgBox "Already at the first record.", vbInformation
    End If

Set DataGrid1.DataSource = rec
HidecarIDColumn

End Sub

Private Sub Command7_Click()
    ' Make sure all required textboxes have values
    If txtplate.Text = "" Or txtbrand.Text = "" Or txtname.Text = "" Or cmbseater.Text = "" Or txtprice.Text = "" Then
        Exit Sub
    End If

    ' Make sure recordset exists and is open
    If rec Is Nothing Then Exit Sub
    If rec.State <> adStateOpen Then Exit Sub
    If rec.BOF And rec.EOF Then Exit Sub

    ' Move to next record
    If rec.AbsolutePosition < rec.RecordCount Then
        rec.MoveNext
        UpdateTextBoxes   ' Populate textboxes with the new current record
    Else
        MsgBox "Already at the last record.", vbInformation
    End If

Set DataGrid1.DataSource = rec
HidecarIDColumn
End Sub



Private Sub Command8_Click()
    ' 1?? Ensure all TextBoxes have values
    If txtplate.Text = "" Or txtbrand.Text = "" Or txtname.Text = "" Or cmbseater.Text = "" Or txtprice.Text = "" Then
        Exit Sub
    End If

    ' 2?? Ensure the recordset exists and is open
    If rec Is Nothing Then Exit Sub
    If rec.State <> adStateOpen Then Exit Sub
    If rec.BOF And rec.EOF Then Exit Sub

    ' 3?? Move to the last record
    rec.MoveLast
    UpdateTextBoxes   ' Populate TextBoxes with the current record

Set DataGrid1.DataSource = rec
HidecarIDColumn
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrHandler

    Dim searchValue As String
    Dim oldFilter As String
    
    searchValue = Trim(Text5.Text)
    oldFilter = rec.Filter  ' Save current filter

    ' If search box is empty, show all records
    If searchValue = "" Then
        rec.Filter = ""
        Set DataGrid1.DataSource = rec
        DataGrid1.Refresh
        Exit Sub
    End If

    ' Convert input to uppercase
    searchValue = UCase(searchValue)

    ' Apply filter (match anywhere in Plate)
    rec.Filter = "Plate LIKE '*" & searchValue & "*'"

    ' Safe check: ensure recordset has at least 1 record
    If rec.EOF And rec.BOF Then
        MsgBox "No record found!", vbInformation
        ' Restore old filter safely
        If Len(oldFilter) > 0 Then
            rec.Filter = oldFilter
        Else
            rec.Filter = ""
        End If
    Else
        ' Only bind DataGrid if records exist
        Set DataGrid1.DataSource = rec
        DataGrid1.Refresh
    End If

    Exit Sub

ErrHandler:
    ' Catch unexpected errors
    rec.Filter = ""
    Set DataGrid1.DataSource = rec
    DataGrid1.Refresh
    Err.Clear
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
    txtbrand.Text = IIf(IsNull(rec!Brand), "", rec!Brand)
    txtname.Text = IIf(IsNull(rec!Name), "", rec!Name)
    comboseat.Text = IIf(IsNull(rec!Seater), "", rec!Seater)
   
    
    ' Show numeric value in TextBox (no $ sign in TextBox)
    If IsNull(rec!Price) Then
    txtprice.Text = ""
Else
    txtprice.Text = CStr(rec!Price)
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
    ' Center form
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    ' Open connection
    Set con = New ADODB.Connection
    con.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\MEQODE\mdb\MasterData.mdb;"
    con.Open

    ' Open recordset
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient
    rec.CursorType = adOpenStatic        ' ? safer for DataGrid
    rec.LockType = adLockOptimistic
    rec.Open "SELECT * FROM vehicles", con

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

    ' Ensure editable for new entry
    isEditing = False
    SetFieldsEditable False
End Sub





Private Sub Text5_KeyPress(KeyAscii As Integer)
    ' Allow Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Get current cursor position
    Dim pos As Integer
    pos = Text5.SelStart

    ' Limit total length to 7 if not replacing selected text
    If Text5.SelLength = 0 And Len(Text5.Text) >= 7 Then
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

Private Sub txtprice_Change()

    Static formatting As Boolean
    If formatting Then Exit Sub
    formatting = True

    Dim raw As String
    Dim num As Double
    Dim decPart As String
    Dim dotPos As Integer
    Dim pesoSign As String

    pesoSign = "Php "  ' Use Php instead of ?

    ' Remove peso sign and commas
    raw = txtprice.Text
    raw = Replace(raw, pesoSign, "")
    raw = Replace(raw, ",", "")

    ' If empty
    If raw = "" Then
        txtprice.Text = ""
        formatting = False
        Exit Sub
    End If

    ' Split decimal
    dotPos = InStr(raw, ".")
    If dotPos > 0 Then
        decPart = Mid(raw, dotPos) ' Keep decimal part
        num = Val(Left(raw, dotPos - 1))
    Else
        num = Val(raw)
        decPart = ""
    End If

    ' LIMIT to 900,000
    If num > 900000 Then num = 900000

    ' Format back with peso sign
    txtprice.Text = pesoSign & Format(num, "#,##0") & decPart

    txtprice.SelStart = Len(txtprice.Text)

    formatting = False

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
