VERSION 5.00
Begin VB.Form frmReplace 
   Caption         =   "Find/Replace"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form3"
   ScaleHeight     =   2640
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   5355
      TabIndex        =   17
      Top             =   0
      Width           =   5775
   End
   Begin VB.CommandButton cmdFindAll 
      Caption         =   "Find All"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkUnescape 
      Caption         =   "Use %xx for hex character values"
      Height          =   240
      Left            =   1035
      TabIndex        =   15
      Top             =   945
      Width           =   2685
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   1350
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find First"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   900
      Width           =   1335
   End
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Selection"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Whole Text"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2250
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label lblSelSize 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Hex"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Char"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Replace"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:   dzzie@yahoo.com
'Site:     http://sandsprite.com

Public SCI As scisimple
Dim lastkey As Integer
Dim lastIndex As Long
Dim lastsearch As String
Dim init As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_SHOWWINDOW = &H40

Friend Sub SetFindText(x As String)
    Text1 = x
    
    'scan text2 (replace text) for numeric extension increment if found..
    On Error Resume Next
    Dim ext, a, c
    
    ext = ""
    x = Trim(Text2)
    a = Len(x)
    
    If a = 0 Then
        'MsgBox "len 0"
        Exit Sub
    End If
    
    'must be at least one numeric char at end
    If Not IsNumeric(Mid(x, a)) Then
        'MsgBox "not numeric """ & Mid(x, a) & """"
        Exit Sub
    End If
    
    'MsgBox 2
    For i = 0 To 10
        c = Mid(x, a - i)
        If Not IsNumeric(c) Then
            Exit For
        Else
            ext = c
        End If
    Next
    
    'MsgBox ext
    If Len(ext) > 0 Then
        ext = CLng(ext)
        Text2 = Mid(x, 1, Len(x) - Len(ext)) & ext + 1
    End If
    
End Sub

Private Sub cmdFind_Click()
    
    On Error Resume Next
    
    If chkUnescape.Value = 1 Then
        f = unescape(Text1)
    Else
        f = Text1
    End If
    
    lastsearch = f
    
    Dim compare As VbCompareMethod
    
    If chkCaseSensitive.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    
    x = InStr(1, SCI.Text, lastsearch, compare)
    If x > 0 Then
        lastIndex = x + 2
        SCI.SelStart = x - 1
        SCI.SelLength = Len(lastsearch)
        SCI.GotoLine SCI.CurrentLine - 1
        SCI.SelectLine
        Me.Caption = "Line: " & SCI.CurrentLine & " CharPos: " & SCI.SelStart
    Else
        lastIndex = 1
    End If
    
End Sub


Public Sub cmdFindAll_Click()
    
    On Error Resume Next
    Dim txt As String
    Dim line As Long
    Dim editorText As String
    Dim firstLine As Long
    
    If Me.Width < 10440 Then Me.Width = 10440
    List1.Clear
    
    If chkUnescape.Value = 1 Then
        f = unescape(Text1)
    Else
        f = Text1
    End If
    
    Dim compare As VbCompareMethod
    
    If chkCaseSensitive.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    
    firstLine = SCI.DirectSCI.GetFirstVisibleLine
    lastIndex = 1
    lastsearch = f
    x = 1
    
    If Len(f) = 0 Then Exit Sub
    
    LockWindowUpdate SCI.sciHWND
    editorText = SCI.Text
    Do While x > 0
    
        x = InStr(lastIndex, editorText, lastsearch, compare)
    
        If x + 2 = lastIndex Or x < 1 Or x >= Len(editorText) Then
            Exit Do
        Else
            lastIndex = x + 2
            SCI.SelStart = x - 1
            SCI.SelLength = Len(lastsearch)
            line = SCI.CurrentLine
            txt = Replace(Trim(SCI.GetLineText(line)), vbTab, Empty)
            txt = Replace(txt, vbCrLf, Empty)
            List1.AddItem (line + 1) & ": " & txt
            
            ' Save some time here.  Since were marking all instances if the same
            ' string is found twice in the same line we don't need to know that.
            ' So once we find it in a line and mark it automaticly jump to the next
            ' line

            SCI.DirectSCI.GotoLine line + 1
            lastIndex = SCI.SelStart
            
        End If
        
    Loop
    
    SCI.DirectSCI.GotoLine firstLine
    LockWindowUpdate 0
    
    'I dont like the auto navigation..disorientating..let the user decide
    'If List1.ListCount >= 0 Then
        'List1.selected(0) = True
        'List1_Click
    'End If
    
    Me.Caption = List1.ListCount & " items found!"
    
End Sub

Private Sub cmdFindNext_Click()
    
    On Error Resume Next
    
    If chkUnescape.Value = 1 Then
        f = unescape(Text1)
    Else
        f = Text1
    End If
    
    If lastsearch <> f Then
        cmdFind_Click
        Exit Sub
    End If
    
    If lastIndex >= Len(SCI.Text) Then
        MsgBox "Reached End of text no more matches", vbInformation
        Exit Sub
    End If
    
    Dim compare As VbCompareMethod
    
    If chkCaseSensitive.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    
    x = InStr(lastIndex, SCI.Text, lastsearch, compare)
    
    If x + 2 = lastIndex Or x < 1 Then
        MsgBox "No more matches found", vbInformation
        Exit Sub
    Else
        lastIndex = x + 2
        SCI.SelStart = x - 1
        SCI.SelLength = Len(lastsearch)
        SCI.GotoLine SCI.CurrentLine
        SCI.SelectLine
        Me.Caption = "Line: " & SCI.CurrentLine & " CharPos: " & SCI.SelStart
    End If
    
    
End Sub

Private Sub cmdReplace_Click()
    
    On Error Resume Next
    
    If chkUnescape.Value = 1 Then
        f = unescape(Text1)
    Else
        f = Text1
    End If
    
    If chkUnescape.Value = 1 Then
        r = unescape(Text2)
    Else
        r = Text2
    End If
    
    Dim compare As VbCompareMethod
    
    If chkCaseSensitive.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    
    Dim curLine As Long
    
    If Option1.Value Then 'whole selection
        curLine = SCI.FirstVisibleLine
        SCI.Text = Replace(SCI.Text, f, r, , , compare)
        SCI.FirstVisibleLine = curLine
    Else
        sl = SCI.SelStart
        nt = Replace(SCI.SelText, f, r, , , compare)
        SCI.SelText = nt
        SCI.SelStart = sl
        SCI.SelLength = Len(nt)
    End If
    
    lblSelSize = "Selection Size: " & Len(SCI.SelText)
    
End Sub

Public Sub LaunchReplaceForm(txtObj As scisimple)
    On Error Resume Next
    Set SCI = txtObj
    SCI.ReplaceFormActive = True
    If Len(txtObj.SelText) > 1 Then
        lblSelSize = "Selection Size: " & Len(txtObj.SelText)
        Text1 = txtObj.SelText
    End If
    If txtObj.ReadOnly Then
        cmdReplace.Enabled = False
        Me.Caption = "Read Only"
    End If
    cmdFindAll.visible = True
    Me.show
    If Not init Then Form_Load
    Form_Resize
End Sub




Private Sub Form_Load()
    init = True
    FormPos Me, True
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_SHOWWINDOW
    If Len(Text1) = 0 Then Text1 = GetMySetting("lastFind")
    Text2 = GetMySetting("lastReplace")
    If GetMySetting("wholeText", "1") = "1" Then Option1.Value = True Else Option2.Value = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    List1.Height = Me.Height - List1.Top - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, True, True
    SaveMySetting "lastFind", Text1
    SaveMySetting "lastReplace", Text2
    SaveMySetting "wholeText", IIf(Option1.Value, "1", "0")
    On Error Resume Next
    SCI.ReplaceFormActive = False
End Sub

Private Function ListSelIndex(lst As ListBox) As Long
    
    On Error GoTo hell
    
    For i = 0 To List1.ListCount
        If List1.selected(i) Then
            ListSelIndex = i
            Exit Function
        End If
    Next

hell:
    ListSelIndex = -1
    
End Function

Private Sub List1_Click()

    On Error Resume Next
    
    Dim tmp As String
    Dim line As Long
    Dim index As Long
    
    index = ListSelIndex(List1)
    
    If index >= 0 Then
        tmp = List1.List(index)
        If InStr(1, tmp, ":") > 0 Then
            line = CLng(Split(tmp, ":")(0))
            SCI.GotoLineCentered line
        End If
    End If
    
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyAll_Click()
    On Error Resume Next
    Dim x As String
    For i = 0 To List1.ListCount
        x = x & List1.List(i) & vbCrLf
    Next
    Clipboard.Clear
    Clipboard.SetText x
    MsgBox Len(x) & " bytes copied", vbInformation
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    lastkey = KeyAscii
End Sub

Private Sub Text3_KeyUp(KeyAscii As Integer, Shift As Integer)
    Dim x As String
    x = Hex(lastkey)
    If Len(x) = 1 Then x = "0" & x
    Text4 = x
    Text3 = Chr(lastkey)
End Sub

Public Function isHexChar(hexValue As String, Optional b As Byte) As Boolean
    On Error Resume Next
    Dim v As Long
    
    
    If Len(hexValue) = 0 Then GoTo nope
    If Len(hexValue) > 2 Then GoTo nope 'expecting hex char code like FF or 90
    
    v = CLng("&h" & hexValue)
    If Err.Number <> 0 Then GoTo nope 'invalid hex code
    
    b = CByte(v)
    If Err.Number <> 0 Then GoTo nope  'shouldnt happen.. > 255 cant be with len() <=2 ?

    isHexChar = True
    
    Exit Function
nope:
    Err.Clear
    isHexChar = False
End Function

Private Function hex_bpush(bAry() As Byte, hexValue As String) As Boolean   'this modifies parent ary object
    On Error Resume Next
    Dim b As Byte
    If Not isHexChar(hexValue, b) Then Exit Function
    bpush bAry, b
    hex_bpush = True
End Function


'this should now be unicode safe on foreign systems..
Function unescape(x) As String '%uxxxx and %xx
    
    'On Error GoTo hell
    
    Dim tmp() As String
    Dim b1 As String, b2 As String
    Dim i As Long
    Dim r() As Byte
    Dim elems As Long
    Dim t
    
    tmp = Split(x, "%")
    
    s_bpush r(), tmp(0) 'any prefix before encoded part..
    
    For i = 1 To UBound(tmp)
        t = tmp(i)
        
        If LCase(VBA.Left(t, 1)) = "u" Then
        
            If Len(t) < 5 Then '%u21 -> %u0021
                t = "u" & String(5 - Len(t), "0") & Mid(t, 2)
            End If

            b1 = Mid(t, 2, 2)
            b2 = Mid(t, 4, 2)
            
            If isHexChar(b1) And isHexChar(b2) Then
                hex_bpush r(), b2
                hex_bpush r(), b1
            Else
                s_bpush r(), "%u" & b1 & b2
            End If
            
            If Len(t) > 5 Then s_bpush r(), Mid(t, 6)
             
        Else
               b1 = Mid(t, 1, 2)
               If Not hex_bpush(r(), b1) Then s_bpush r(), "%" & b1
               If Len(t) > 2 Then s_bpush r(), Mid(t, 3)
        End If
        
    Next
            
hell:
    unescape = StrConv(r(), vbUnicode, LANG_US)
     
     If Err.Number <> 0 Then
        MsgBox "Error in unescape: " & Err.Description
     End If
     
End Function

Private Sub s_bpush(bAry() As Byte, sValue As String)
    Dim tmp() As Byte
    Dim i As Long
    tmp() = StrConv(sValue, vbFromUnicode, LANG_US)
    For i = 0 To UBound(tmp)
        bpush bAry, tmp(i)
    Next
End Sub

Private Sub bpush(bAry() As Byte, b As Byte) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    
    x = UBound(bAry) '<-throws Error If Not initalized
    ReDim Preserve bAry(UBound(bAry) + 1)
    bAry(UBound(bAry)) = b
    
    Exit Sub

init:
    ReDim bAry(0)
    bAry(0) = b
    
End Sub
