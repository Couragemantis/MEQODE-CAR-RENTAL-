VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmcustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Registration Management System"
   ClientHeight    =   7050
   ClientLeft      =   105
   ClientTop       =   450
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
      Left            =   10200
      Top             =   6120
   End
   Begin VB.TextBox txtln 
      Height          =   405
      Left            =   2040
      TabIndex        =   19
      Top             =   1440
      Width           =   4935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   7440
      TabIndex        =   18
      Top             =   2040
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   134086657
      CurrentDate     =   46078
   End
   Begin VB.ComboBox combodt 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search License Number"
      Height          =   855
      Left            =   7440
      TabIndex        =   11
      Top             =   1200
      Width           =   6375
      Begin VB.CommandButton Command9 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   120
         TabIndex        =   12
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
      Left            =   1320
      TabIndex        =   6
      Top             =   4920
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "&Archive"
         Height          =   495
         Left            =   3360
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Update"
         Height          =   495
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton ClearInputs 
         Caption         =   "&New"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox txtcn 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   4935
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Date"
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
      Left            =   75
      TabIndex        =   16
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   13680
      Picture         =   "frmcustomer.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Customer"
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
      TabIndex        =   14
      Top             =   0
      Width           =   14535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Driver Type"
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
      Left            =   75
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
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
      Left            =   75
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   75
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "License Number"
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
      Left            =   75
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   -120
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "frmcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private db As clsDB
Private rec As ADODB.Recordset
Private isEditing As Boolean
Private bIsTyping As Boolean
Private Sub txtln_GotFocus()
    bIsTyping = True
End Sub

Private Sub txtln_LostFocus()
    bIsTyping = False
End Sub

Private Sub txtname_GotFocus()
    bIsTyping = True
End Sub

Private Sub txtname_LostFocus()
    bIsTyping = False
End Sub

Private Sub txtcn_GotFocus()
    bIsTyping = True
End Sub

Private Sub txtcn_LostFocus()
    bIsTyping = False
End Sub
Private Sub RefreshGridNoSelection()

    rec.Requery
    Set DataGrid1.DataSource = rec
    HidecusIDColumn

    If Not (rec.BOF And rec.EOF) Then
        rec.MoveLast
        rec.MoveNext
    End If

End Sub
Private Sub HidecusIDColumn()
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
    cmbseater.Text = rec!Seater
    txtprice.Text = rec!price
End Sub


Private Sub SetFieldsEditable(Optional ByVal editing As Boolean = False, Optional ByVal adding As Boolean = False)
    ' Text1 (Student ID) behavior:
    ' - Locked when editing an existing record
    ' - Editable when adding a new record
    If editing And Not adding Then
        txtln.Locked = True
    Else
        txtln.Locked = False
    End If

    ' Other fields
    txtln.Locked = False
    txtname.Locked = False
    txtcn.Locked = False
End Sub

Private Sub ClearInputs_Click()
    txtln.Text = ""
    txtname.Text = ""
    txtcn.Text = ""

    combodt.ListIndex = -1
    
    RefreshGridNoSelection
End Sub

Private Sub cmdAdd_Click()
    
    ' ==============================
    ' Validate Required Fields
    ' ==============================
    If Trim(txtln.Text) = "" Or _
       Trim(txtname.Text) = "" Or _
       Trim(txtcn.Text) = "" Or _
       combodt.ListIndex = -1 Then
       
        MsgBox "Please fill in all required fields!", vbExclamation
        Exit Sub
    End If

    ' ==============================
    ' Validate License Format (1 Letter + 10 Digits)
    ' ==============================
    If Not txtln.Text Like "[A-Za-z]##########" Then
        MsgBox "License must be 1 letter followed by 10 digits!", vbExclamation
        txtln.SetFocus
        Exit Sub
    End If

    ' ==============================
    ' Validate Contact Number (exactly 11 digits)
    ' ==============================
    Dim ContactNo As String
    ContactNo = Trim(txtcn.Text)
    If ContactNo <> "" And Not ContactNo Like "09#########" Then
        MsgBox "Invalid Contact Number.", vbExclamation
        txtcn.SetFocus
        Exit Sub
    End If

    ' ==============================
    ' Check Duplicate License
    ' ==============================
    Dim rsCheck As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset

    rsCheck.Open "SELECT License FROM customer WHERE License = '" & _
                 db.SafeText(txtln.Text) & "'", _
                 db.con, adOpenForwardOnly, adLockReadOnly
                 
    If Not rsCheck.EOF Then
        MsgBox "License already exists!", vbExclamation
        rsCheck.Close
        Set rsCheck = Nothing
        Exit Sub
    End If

    rsCheck.Close
    Set rsCheck = Nothing

    ' ==============================
    ' Store values for message
    ' ==============================
    Dim addedLicense As String, addedName As String
    addedLicense = txtln.Text
    addedName = txtname.Text

    ' ==============================
    ' Add New Customer Record
    ' ==============================
    rec.AddNew
    rec!License = addedLicense
    rec!Name = addedName
    rec!Contact = txtcn.Text
    rec!Type = combodt.Text
    rec!Expiration = Format(DTPicker1.Value, "mm/dd/yyyy")
    rec.Update

    ' ==============================
    ' Refresh DataGrid
    ' ==============================
    RefreshGridNoSelection

    ' ==============================
    ' Clear input fields
    ' ==============================
    txtln.Text = ""
    txtname.Text = ""
    txtcn.Text = ""
    combodt.ListIndex = -1

    txtln.SetFocus
    isEditing = False
    SetFieldsEditable False

    ' ==============================
    ' Show success message with License and Name
    ' ==============================
    MsgBox "Customer added successfully: " & addedLicense & " - " & addedName, vbInformation

End Sub
Private Sub cmdEdit_Click()
    
    ' ==============================
    ' Check if record is selected
    ' ==============================
    If rec.EOF Or rec.BOF Then
        MsgBox "Please select a record to edit!", vbExclamation
        Exit Sub
    End If

    ' ==============================
    ' Validate required fields
    ' ==============================
    If Trim(txtln.Text) = "" Or _
       Trim(txtname.Text) = "" Or _
       Trim(txtcn.Text) = "" Or _
       combodt.ListIndex = -1 Then
       
        MsgBox "Please fill in all required fields!", vbExclamation
        Exit Sub
    End If

    ' ==============================
    ' Validate License format: 1 letter + 10 digits
    ' ==============================
    If Not txtln.Text Like "[A-Za-z]##########" Then
        MsgBox "License must be 1 letter followed by 10 digits!", vbExclamation
        txtln.SetFocus
        Exit Sub
    End If

    ' ==============================
    ' Validate Name length
    ' ==============================
    If Len(txtname.Text) > 50 Then
        MsgBox "Name cannot exceed 50 characters!", vbExclamation
        txtname.SetFocus
        Exit Sub
    End If

    ' ==============================
    ' Validate Contact (numeric + exactly 11 digits)
    ' ==============================
    If Not IsNumeric(txtcn.Text) Or Len(txtcn.Text) <> 11 Then
        MsgBox "Contact must be numeric and exactly 11 digits!", vbExclamation
        txtcn.SetFocus
        Exit Sub
    End If

    ' ==============================
    ' Check if any changes were made
    ' ==============================
    If txtln.Text = rec!License And _
       txtname.Text = rec!Name And _
       txtcn.Text = rec!Contact And _
       combodt.Text = rec!Type And _
       Format(DTPicker1.Value, "mm/dd/yyyy") = Format(rec!Expiration, "mm/dd/yyyy") Then
       
        MsgBox "No changes detected!", vbInformation
        Exit Sub
    End If

    ' ==============================
    ' Confirm save
    ' ==============================
    If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Confirm Edit") = vbYes Then

        ' Update record
        rec!License = txtln.Text
        rec!Name = txtname.Text
        rec!Contact = txtcn.Text
        rec!Type = combodt.Text
        rec!Expiration = Format(DTPicker1.Value, "mm/dd/yyyy")
        rec.Update

        ' Refresh DataGrid
        RefreshGridNoSelection
        ' Sync TextBoxes


        MsgBox "Customer record updated successfully!", vbInformation

        ' Clear input fields
        txtln.Text = ""
        txtname.Text = ""
        txtcn.Text = ""
        combodt.ListIndex = -1




        txtln.SetFocus
        isEditing = False
        SetFieldsEditable False

    End If

End Sub


Private Sub Command4_Click()

    ' ==========================
    ' Ensure a record is selected
    ' ==========================
    If rec.EOF Or rec.BOF Then
        MsgBox "Please select a record to archive!", vbExclamation
        Exit Sub
    End If

    ' ==========================
    ' Confirm action
    ' ==========================
    If MsgBox("Are you sure you want to archive this record?", vbYesNo + vbQuestion, "Confirm Archive") = vbYes Then
        
        Dim conArchive As ADODB.Connection
        Dim rsArchive As ADODB.Recordset
        
        ' ==========================
        ' Open archive database
        ' ==========================
        Set conArchive = New ADODB.Connection
        conArchive.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\Archives\customerarchive\customerarchives.mdb;"
        conArchive.Open
        
        Set rsArchive = New ADODB.Recordset
        rsArchive.CursorType = adOpenDynamic
        rsArchive.LockType = adLockOptimistic
        rsArchive.Open "SELECT * FROM customerarchives", conArchive
        
        ' ==========================
        ' Copy selected record to archive
        ' ==========================
        rsArchive.AddNew
        rsArchive!License = rec!License
        rsArchive!Name = rec!Name
        rsArchive!Contact = rec!Contact
        rsArchive!Type = rec!Type
        rsArchive!Expiration = rec!Expiration
        rsArchive.Update
        
        rsArchive.Close
        conArchive.Close
        Set rsArchive = Nothing
        Set conArchive = Nothing

        ' ==========================
        ' Delete from main database
        ' ==========================
        rec.Delete
        rec.Update
        rec.Requery
        rec.MoveLast
        rec.MoveNext   ' forces EOF (no selection)

        ' Refresh DataGrid
        Set DataGrid1.DataSource = rec
        HidecusIDColumn
        
        ' ==========================
        ' Clear form
        ' ==========================
        txtln.Text = ""
        txtname.Text = ""
        txtcn.Text = ""
        combodt.ListIndex = -1
        txtln.SetFocus
        
        isEditing = False
        SetFieldsEditable False
        
        MsgBox "Record archived successfully!", vbInformation
        
    End If  ' <-- Correct End If for the MsgBox confirmation

End Sub
Private Sub Command9_Click()
    On Error GoTo ErrHandler  ' Catch unexpected errors

    Dim searchValue As String
    Dim oldFilter As String
    
    searchValue = Trim(Text5.Text)
    
    ' Save current filter
    oldFilter = rec.Filter
    
    ' If textbox is empty, show all records
    If searchValue = "" Then
        rec.Filter = ""
        Set DataGrid1.DataSource = rec
        DataGrid1.Refresh
        Exit Sub
    End If
    
    ' Convert to uppercase
    searchValue = UCase(searchValue)
    
    ' Ensure only max 11 characters
    If Len(searchValue) > 11 Then searchValue = Left(searchValue, 11)
    
    ' Try applying filter safely
    On Error Resume Next
    rec.Filter = "License LIKE '*" & searchValue & "*'"
    
    ' If an error occurs or no records found, restore filter
    If Err.Number <> 0 Or rec.EOF Then
        Err.Clear
        rec.Filter = oldFilter
        Set DataGrid1.DataSource = rec
        DataGrid1.Refresh
        MsgBox "No record found or invalid input!", vbInformation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Update DataGrid
    Set DataGrid1.DataSource = rec
    DataGrid1.Refresh
    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred. Showing all records.", vbExclamation
    rec.Filter = ""
    Set DataGrid1.DataSource = rec
    DataGrid1.Refresh
End Sub

Private Sub DataGrid1_Click()

    ' Exit if recordset invalid
    If rec Is Nothing Then Exit Sub
    If rec.EOF Or rec.BOF Then Exit Sub

    ' Move to selected row safely
    On Error Resume Next
    rec.Bookmark = DataGrid1.Bookmark
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If rec.EOF Or rec.BOF Then Exit Sub

    ' ==========================
    ' Populate textboxes
    ' ==========================
    txtln.Text = IIf(IsNull(rec!License), "", rec!License)
    txtname.Text = IIf(IsNull(rec!Name), "", rec!Name)
    txtcn.Text = IIf(IsNull(rec!Contact), "", rec!Contact)

    ' ==========================
    ' Populate ComboBox
    ' ==========================
    Dim typeValue As String
    typeValue = IIf(IsNull(rec!Type), "", rec!Type)
    
    ' Set ListIndex based on value in combodt
    Select Case typeValue
        Case "Non - Professional"
            combodt.ListIndex = 0
        Case "Professional"
            combodt.ListIndex = 1
        Case Else
            combodt.ListIndex = -1
    End Select

    ' Update connected textbox
    DTPicker1.Value = Value

    ' ==========================
    ' Populate expiration date
    ' ==========================
    If Not IsNull(rec!Expiration) Then
        DTPicker1.Value = rec!Expiration
        
    Else
        DTPicker1.Value = "01/01/2000"
    End If

    ' Set editable state
    SetFieldsEditable isEditing

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
        KeyCode = 0
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub Form_Load()
    ' Form centering
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    ' Initialize DB (form-level db)
    Set db = New clsDB
    db.OpenDB

    ' Open customer recordset
    Set rec = New ADODB.Recordset
    With rec
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM customer", db.con  ' now works
    End With

    ' Bind to DataGrid
    Set DataGrid1.DataSource = rec
    HidecusIDColumn

    ' Initialize ComboBox
    combodt.Clear
    combodt.AddItem "Non - Professional"
    combodt.AddItem "Professional"
    combodt.ListIndex = -1

    ' DatePicker
    DTPicker1.Format = dtpCustom
    DTPicker1.CustomFormat = "MM/dd/yyyy"

    ' Edit state
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
    Dim pos As Integer
    Dim selLength As Integer

    ' Exit if Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Get current selection
    selLength = Text5.selLength

    ' Calculate the new length if replacing selection
    pos = Len(Text5.Text) - selLength + 1

    ' Limit maximum characters to 11
    If pos > 11 Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If

    ' Validate character by position
    Select Case pos
        Case 1
            ' First character must be a letter
            If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or _
                    (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = 0
                Beep
            Else
                ' Convert to uppercase automatically
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        Case 2 To 11
            ' Remaining characters must be numbers
            If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
                KeyAscii = 0
                Beep
            End If
    End Select
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrAutoRefresh_Timer()
    If rec Is Nothing Then Exit Sub
    If rec.State <> adStateOpen Then Exit Sub

    ' Skip refresh while typing
    If bIsTyping Then Exit Sub

    If db.HasTableChanged("customer") Then
        RefreshGridNoSelection
    End If
End Sub
Private Sub txtcn_KeyPress(KeyAscii As Integer)

    ' Allow Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Allow only numbers 0-9
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If

    ' Handle selection replacement
    With txtcn
        If .selLength > 0 Then
            ' Calculate new length after replacing selection
            If Len(.Text) - .selLength + 1 > 11 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            ' Replace selection with typed key
            .Text = Left(.Text, .SelStart) & Chr(KeyAscii) & Mid(.Text, .SelStart + .selLength + 1)
            ' Move cursor after inserted digit
            .SelStart = .SelStart + 1
            ' Prevent default insertion
            KeyAscii = 0
            Exit Sub
        End If
    End With

    ' Limit length to 11 characters if no selection
    If Len(txtcn.Text) >= 11 Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    
    

End Sub

            

Private Sub txtln_KeyDown(KeyCode As Integer, Shift As Integer)
 
    ' Cancel DELETE key
    If KeyCode = 46 Then
        KeyCode = 0
        Beep
    End If
End Sub

Private Sub txtln_KeyPress(KeyAscii As Integer)
    

    ' Determine current position
    Dim pos As Integer
    pos = Len(txtln.Text) + 1

    ' Always allow first character rules
    Select Case pos
        Case 1
            If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = 0
                Beep
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii))) ' Auto uppercase
            End If
        Case 2 To 11
            If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
                KeyAscii = 0
                Beep
            End If
        Case Else
            KeyAscii = 0
            Beep
    End Select

    ' Lock input automatically when 11 characters typed
    
End Sub
Private Sub txtname_KeyPress(KeyAscii As Integer)

    ' Allow Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Allow letters (A-Z, a-z), dot (.), and space
    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Or _
       KeyAscii = 46 Or _
       KeyAscii = 32 Then

        ' Replace selected text with typed key
        With txtname
            If .selLength > 0 Then
                ' Replace selection with typed key
                .Text = Left(.Text, .SelStart) & Chr(KeyAscii) & Mid(.Text, .SelStart + .selLength + 1)
                ' Move cursor after the typed key
                .SelStart = .SelStart + 1
                ' Prevent default insertion
                KeyAscii = 0
            End If
        End With

    Else
        ' Invalid key
        KeyAscii = 0
        Beep
    End If

End Sub


