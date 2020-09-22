Attribute VB_Name = "Module1"

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public ret As String
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
 
 Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
 End Type
 
 Public Const BDR_SUNKENOUTER = &H2
 Public Const BF_LEFT = &H1
 Public Const BF_TOP = &H2
 Public Const BF_RIGHT = &H4
 Public Const BF_BOTTOM = &H8
 Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
 Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

Public Sub DrawSunkenFrame(picbox)
  Dim PicRect As RECT
  With picbox
    .Cls
    PicRect.Left = .ScaleLeft
    PicRect.Top = .ScaleTop
    PicRect.Right = .ScaleWidth
    PicRect.Bottom = .ScaleHeight
  End With
DrawEdge picbox.hDC, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
If picbox.AutoRedraw Then picbox.Refresh
End Sub

Function SaveMetaFile()
  Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "META Files (*.met)" + Chr$(0) + "*.met" + Chr$(0) + "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path ' CurDir
        ofn.lpstrTitle = "Save META Tag File"
        ofn.flags = 0
        Dim a
        a = GetSaveFileName(ofn)

        If (a) Then
                SaveMetaFile = Trim$(ofn.lpstrFile)
        Else
                SaveMetaFile = "cancel"
        End If
End Function

Function OpenMetaFile()
  Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "META Files (*.met)" + Chr$(0) + "*.met" + Chr$(0) + "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path ' cur dir
        ofn.lpstrTitle = "Open META Tag File"
        ofn.flags = 0
        Dim a
        a = GetOpenFileName(ofn)

        If (a) Then
                OpenMetaFile = Trim$(ofn.lpstrFile)
        Else
                OpenMetaFile = "cancel"
        End If
End Function

Public Sub WriteINI(FileName As String, Section, Key As String, Text As String)
WritePrivateProfileString Section, Key, Text, FileName
End Sub

Public Function ReadINI(FileName As String, Section, Key As String)
ret = Space$(255)
retlen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), FileName)
ret = Left$(ret, retlen)
ReadINI = ret
End Function
