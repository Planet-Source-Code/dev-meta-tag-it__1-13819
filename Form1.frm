VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meta Tag IT"
   ClientHeight    =   3375
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   5745
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox tabn 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   3
      Left            =   240
      ScaleHeight     =   2655
      ScaleWidth      =   5295
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   42
         Text            =   "Form1.frx":030A
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Help"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Copy"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   40
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "HTML Code:  "
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox tabn 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   2
      Left            =   240
      ScaleHeight     =   2655
      ScaleWidth      =   5295
      TabIndex        =   37
      Top             =   480
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   52
         Text            =   "http://www.geocities.com/vbfortress"
         Top             =   2280
         Width           =   4935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Refresh to new location"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Do not cache page"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Page Rating:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   47
         Top             =   960
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Generator:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   45
         Top             =   600
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Author:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Misc Tags:  "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox tabn 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   1
      Left            =   240
      ScaleHeight     =   2655
      ScaleWidth      =   5295
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text2 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "The Description should contain some of the keywords you used."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   4695
      End
   End
   Begin VB.PictureBox tabn 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   0
      Left            =   240
      ScaleHeight     =   2655
      ScaleWidth      =   5295
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   4080
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   4080
         TabIndex        =   30
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   4080
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   4080
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   4080
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   21
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   20
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Extra Keywords:  "
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblnum 
         Caption         =   "15."
         Height          =   255
         Index           =   14
         Left            =   3840
         TabIndex        =   27
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "14."
         Height          =   255
         Index           =   13
         Left            =   3840
         TabIndex        =   26
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "13."
         Height          =   255
         Index           =   12
         Left            =   3840
         TabIndex        =   25
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "12."
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   24
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "11."
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "10."
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   17
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "9."
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   16
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "8."
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   15
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "7."
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "6."
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "5."
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "4."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "3."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "2."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblnum 
         Caption         =   "1."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Keywords"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Site Keywords"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Description"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Site Description"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Misc META Tags"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "HTML Code"
            Object.Tag             =   ""
            Object.ToolTipText     =   "HTML Code"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuload 
         Caption         =   "&Load Metatag File"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save Metatag File"
         Shortcut        =   ^S
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucode_copy 
         Caption         =   "Copy Code"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuclearall 
         Caption         =   "Clear ALL"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuaboutmnu 
      Caption         =   "&About"
      Begin VB.Menu mnuabout 
         Caption         =   "&About "
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub MakeTAG()
Dim enterk, keywords, kwz1: enterk = Chr(13) & Chr(10)

For i = 0 To 14
If Text1(i) <> "" Then
kwz1 = kwz1 + Text1(i) + ", "
End If
Next i

keywords = kwz1 + Text1(15)
Text3 = "<META NAME=" & Chr(34) & "Keywords" & Chr(34) & " CONTENT=" & Chr(34) & keywords & Chr(34) & ">"
Text3 = Text3 & enterk & "<META NAME=" & Chr(34) & "Description" & Chr(34) & " CONTENT=" & Chr(34) & Text2 & Chr(34) & ">"

If Check1(0).Value = 1 Then Text3 = Text3 & enterk & "<META NAME=" & Chr(34) & "AUTHOR" & Chr(34) & " CONTENT=" & Chr(34) & Text4(0) & Chr(34) & ">"
If Check1(1).Value = 1 Then Text3 = Text3 & enterk & "<META NAME=" & Chr(34) & "GENERATOR" & Chr(34) & " CONTENT=" & Chr(34) & Text4(1) & Chr(34) & ">"
If Check3(0).Value = 1 Then Text3 = Text3 & enterk & "<META NAME=" & Chr(34) & "RATING" & Chr(34) & " CONTENT=" & Chr(34) & Combo1 & Chr(34) & ">"
If Check3(1).Value = 1 Then Text3 = Text3 & enterk & "<META HTTP-EQUIV=Pragma CONTENT=NO-CACHE>"
If Check3(2).Value = 1 Then Text3 = Text3 & enterk & "<META HTTP-EQUIV=Refresh CONTENT=" & Chr(34) & "10; URL=" & Text4(2) & Chr(34) & ">"
End Sub
Sub unloadIT()
    Text2 = ""
    Text3 = ""
    Unload aboutbox
    Set aboutbox = Nothing

    Unload Me
    Set Form1 = Nothing
End Sub



Private Sub Check1_Click(Index As Integer)
    Dim x
    x = Check1(Index).Index
    If Check1(Index).Value = 1 Then
    Text4(x).Enabled = True
    Text4(x).BackColor = vbWhite
    Else
    Text4(x).Enabled = False
    Text4(x).BackColor = vbButtonFace
    End If
End Sub



Private Sub Check3_Click(Index As Integer)
    If Check3(Index).Index = 0 Then
    If Check3(0).Value = 1 Then
    Combo1.Enabled = True
    Combo1.BackColor = vbWhite
    Else
    Combo1.Enabled = False
    Combo1.BackColor = vbButtonFace
    End If
    End If
    
    If Check3(Index).Index = 2 Then
    If Check3(2).Value = 1 Then
    Text4(2).Enabled = True
    Text4(2).BackColor = vbWhite
    Else
    Text4(2).Enabled = False
    Text4(2).BackColor = vbButtonFace
    End If
    End If
End Sub

Private Sub Command2_Click(Index As Integer): lne = Chr(13) & Chr(10)
    Select Case Command2(Index).Index
    Case 0
        MsgBox "Once you have entered the data for you meta tags, copy the source code and paste them inbetween the head tags on your site." & lne & lne & "EX:" & lne & lne & "   <HEAD>" & lne & "   META TAGS here" & lne & "   </HEAD>"
    Case 1
        Clipboard.Clear: Clipboard.SetText Text3
    End Select
End Sub

Private Sub Form_Load()
    Combo1.AddItem "General"
    Combo1.AddItem "Adult"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer): unloadIT: End Sub
Private Sub mnuabout_Click()
    aboutbox.Show vbModal
End Sub
Private Sub mnuclearall_Click()
    Dim i
    
    For i = 0 To 15
        Text1(i).Text = ""
    Next i
    
    Text2 = ""
    Text3 = ""
    Text4(0) = ""
    Text4(1) = ""
    Text4(2) = ""
End Sub
Private Sub mnuclose_Click()
    unloadIT
End Sub
Private Sub mnucode_copy_Click()
    Clipboard.Clear
    Clipboard.SetText Text3
End Sub
Private Sub mnuload_Click()
        Dim x$, i
    x$ = OpenMetaFile
    If x$ = "cancel" Then: Exit Sub
    If x$ <> "" Then
      For i = 0 To 14
       Text1(i) = ReadINI(x$, "META Tag IT", "KW" & (i))
      Next i
    End If
    
    Text1(15) = ReadINI(x$, "META Tag IT", "Extra KW")
    Text2 = ReadINI(x$, "META Tag IT", "Description")
    
    p1 = ReadINI(x$, "META Tag IT", "AUTHOR")
    If p1 <> "" Then Text4(0).Text = p1: Check1(0).Value = 1: Text4(0).Enabled = True: Text4(0).BackColor = vbWhite
    p1 = ReadINI(x$, "META Tag IT", "GENERATOR")
    If p1 <> "" Then Text4(1).Text = p1: Check1(1).Value = 1: Text4(1).Enabled = True: Text4(1).BackColor = vbWhite
    p1 = ReadINI(x$, "META Tag IT", "RATING")
    If p1 <> "" Then Combo1.ListIndex = p1: Check3(0).Value = 1: Combo1.Enabled = True: Combo1.BackColor = vbWhite
    p1 = ReadINI(x$, "META Tag IT", "NOCACHE")
    If p1 <> "" Then Check3(1).Value = 1
    p1 = ReadINI(x$, "META Tag IT", "REFRESH")
    If p1 <> "" Then Text4(2).Text = p1: Check3(2).Value = 1: Text4(2).Enabled = True: Text4(2).BackColor = vbWhite
End Sub
Private Sub mnusave_Click()
    Dim x$, i
    x$ = SaveMetaFile
    If x$ = "cancel" Then: Exit Sub
    
    If x$ <> "" Then
      Kill x$
      For i = 0 To 14
       WriteINI x$, "META Tag IT", "KW" & (i), Text1(i)
      Next i
    End If
    
    WriteINI x$, "META Tag IT", "Extra KW", Text1(15)
    WriteINI x$, "META Tag IT", "Description", Text2
    
    If Check1(0).Value = 1 Then WriteINI x$, "META Tag IT", "AUTHOR", Text4(0)
    If Check1(1).Value = 1 Then WriteINI x$, "META Tag IT", "GENERATOR", Text4(1)
    If Check3(0).Value = 1 Then WriteINI x$, "META Tag IT", "RATING", Combo1.ListIndex
    If Check3(1).Value = 1 Then WriteINI x$, "META Tag IT", "NOCACHE", "Y"
    If Check3(2).Value = 1 Then WriteINI x$, "META Tag IT", "REFRESH", Text4(2)
End Sub

Private Sub TabStrip1_Click()
    Dim i: Static CurTab
    If CurTab = Empty Then CurTab = 1
    CurTab = TabStrip1.SelectedItem.Index
    
    For i = 0 To 3
    tabn(i).Visible = False
    Next i
    If CurTab = 1 Then ' KW's
    tabn(0).Visible = True
    ElseIf CurTab = 2 Then ' Description
    tabn(1).Visible = True
    ElseIf CurTab = 3 Then
    tabn(2).Visible = True
    
    ElseIf CurTab = 4 Then ' Kode
    Call MakeTAG
    tabn(3).Visible = True
    End If
End Sub
