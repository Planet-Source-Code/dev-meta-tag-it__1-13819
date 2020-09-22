VERSION 5.00
Begin VB.Form aboutbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About META Tag It"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   430
      Left            =   1200
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   9
      Top             =   3480
      Width           =   1517
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3735
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "http://www.geocities.com/vbfortress"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "the.dev@btinternet.com"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "META TAG  IT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1695
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "META TAG  IT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1695
      Index           =   1
      Left            =   340
      TabIndex        =   1
      Top             =   225
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "META TAG  IT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Index           =   0
      Left            =   380
      TabIndex        =   0
      Top             =   260
      Width           =   3255
   End
End
Attribute VB_Name = "aboutbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Label1(1).Caption = "META Tag It 1.2 Copyright © Dev 1998 - 99"
DrawSunkenFrame pic1
End Sub

