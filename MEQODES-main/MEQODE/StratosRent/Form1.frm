VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VELOCITAS"
   ClientHeight    =   10365
   ClientLeft      =   -4050
   ClientTop       =   300
   ClientWidth     =   23010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   23010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "book"
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "car"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "customer"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Control"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10140
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000A&
      Height          =   735
      Left            =   -120
      Top             =   0
      Width           =   23175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   10695
      Left            =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmcustomer.Show vbModal
End Sub

Private Sub Command2_Click()
frmcar.Show vbModal
End Sub
