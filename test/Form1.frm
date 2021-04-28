VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Begin VB.Form d 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "about"
      Height          =   465
      Left            =   8280
      TabIndex        =   9
      Top             =   585
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   225
      TabIndex        =   8
      Top             =   5535
      Width           =   9420
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   8775
      TabIndex        =   6
      Top             =   135
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove Marker"
      Height          =   420
      Left            =   6750
      TabIndex        =   5
      Top             =   135
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   180
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      Caption         =   "find"
      Height          =   510
      Left            =   3645
      TabIndex        =   3
      Top             =   90
      Width           =   1725
   End
   Begin sci2.SciSimple SciSimple1 
      Height          =   4200
      Left            =   240
      TabIndex        =   2
      Top             =   1140
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7408
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   1710
      TabIndex        =   1
      Top             =   135
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "multiline calltip test list. xx. or yy. for scrolling"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   660
      Width           =   6855
   End
End
Attribute VB_Name = "d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const SCI_SETLEXERLANGUAGE = 4006
 Const SCI_GETLEXERLANGUAGE = 4012
Const SCI_SETSTYLEBITS = 2090
Const SCLEX_ERRORLIST = 10
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim isLocked As Boolean

Dim multi As New CMuiltiLineTip

 

Private Sub cmdAbout_Click()
    SciSimple1.ShowAbout
End Sub

Private Sub Command1_Click()
     
     Dim x As String
     
     With SciSimple1
        .LoadFile "C:\Documents and Settings\david\Desktop\scivb_overflow.js"
'        x = .Text
'         For i = 0 To 6
'             x = x & x
'         Next
'         .Text = x
         
         MsgBox .DirectSCI.GetLineCount
         MsgBox .GetLineText(0)
         'MsgBox "Loaded size= " & FileSize(Len(x))
     End With
     
     
     'MsgBox x
End Sub

Private Sub Command2_Click()
    
'    Dim a As Long, b As Long
'    a = Timer
'    'a = GetTickCount
'    SciSimple1.ExportToHTML "C:\out.html"
'    b = Timer
'    'b = GetTickCount
'
'    MsgBox b - a & " seconds"
    
    Dim f As New d
    f.Show
    f.Move Me.Left + 1000, Me.Top + 1000
    
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Public Function FileSize(bytes As Long) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    'fsize = FileLen(fPath)
    fsize = bytes
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function


Private Sub Command3_Click()
    SciSimple1.ShowFindReplace
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    SciSimple1.DeleteMarker CLng(Text1)
End Sub

Private Sub Command5_Click()
    'isLocked = Not isLocked
    'SciSimple1.LockEditor isLocked
    SciSimple1.ShowFindReplace
End Sub

Private Sub Form_Load()

'    Dim pth As String
'    pth = "C:\Documents and Settings\david\Desktop\scivb\highlighters\java.bin"
'    LoadHighlighter pth
'
'    For i = 0 To 127
'        Highlighter.StyleFont(i) = "Courier New"
'        Highlighter.StyleSize(i) = 12
'    Next
'
'    SaveHighlighter pth
    
    Dim r As Boolean
    
    Set multi.Sci = SciSimple1
    
    With SciSimple1
        r = .LoadHighlighter("D:\_code\scivb2\sql.hilighter")
        List1.AddItem "Load sql: " & r
        
        r = .SetHighlighter("SQL")
         List1.AddItem "Set highlighter sql " & r
         
        .codePage = SC_CP_UTF8
        .WordWrap = True
        .ShowIndentationGuide = True
        .Folding = True
        '.Text = Replace("this func_01 v1 is a simple testaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\nIf it were for real it would not be a test\nBut you already knew that i know\nif(a){\n  alert(2)\n}\n\nadd a dot after this: fsonow type controlh after fso.app\n\n\n\n\ng\ng\ng\ng\njkfdljsfkl\nkdjskldjsl\ndjklfjkds\nhfjdfhjd\nfjkdljfk\njfkdlsjfkl\n", "\n", vbCrLf)
        '.Text = .Text & .Text & .Text & .Text & .Text & .Text
        .Text = "  d"
        .LoadCallTips App.Path & "\..\js_api.txt"
        .AddCallTip "appendfile(blah,blah)"
        .AddCallTip "func_01(test)"
        .ShowIndentationGuide = True
        '.AutoCompleteOnCTRLSpace = False
        '.EnableArrowKeys
        '.SetFocusSci
        
        Const SC_MARK_CIRCLE = 0
        .DirectSCI.MarkerDefine 2, SC_MARK_CIRCLE
        .DirectSCI.MarkerSetFore 2, vbRed 'set breakpoint color
        .DirectSCI.MarkerSetBack 2, vbRed
         
         If Not .LoadFile(App.Path & "\test.js") Then
            .Text = Replace("a = 1; a++; a++; alert(a);", ";", ";" & vbCrLf)
         End If
         
         
    End With
    
    
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    SciSimple1.Width = Me.Width - SciSimple1.Left - 240
    SciSimple1.Height = Me.Height - SciSimple1.Top - 600
End Sub

Private Sub SciSimple1_AutoCompleteEvent(className As String)
    
    Dim prevWord As String
    prevWord = SciSimple1.PreviousWord()

    Debug.Print "AutoCompleteEvent: ClassName: " & className & " Prevword: " & prevWord
        
    'scintinella is smart enough to autoscroll the autocomplete list to the partial match of the curWord :)
    'so fso.app CTRL+H will send us curword=app prevword=fso and sci will scroll to appendfile at top of list.
    
    If className = "tb" Or prevWord = "tb" Then
        SciSimple1.ShowAutoComplete "Save2Clipboard GetClipboard t eval unescape alert Hexdump WriteFile ReadFile HexString2Bytes Disasm pad EscapeHexString GetStream CRC getPageNumWords GetPageNthWord"
    End If
    
    If className = "fso" Or prevWord = "fso" Then
        SciSimple1.ShowAutoComplete "readfile writefile appendfile fileexists deletefile"
    ElseIf className = "ida" Or prevWord = "ida" Then
        'do i want to break these up into smaller chunks for intellisense?
        SciSimple1.ShowAutoComplete "imagebase() loadedfile() jump patchbyte originalbyte readbyte inttohex refresh() " & _
                               "numfuncs() functionstart functionend functionname getasm instsize xrefsto " & _
                               "xrefsfrom undefine getname jumprva screenea() funccount() find " & _
                               "hideea showea hideblock showblock removename setname makecode " & _
                               "getcomment addcomment addcodexref adddataxref delcodexref deldataxref " & _
                               "funcindexfromva funcvabyname nextea prevea patchstring makestr makeunk " & _
                               "renamefunc decompile"
                               
    ElseIf className = "list" Or prevWord = "list" Then
        SciSimple1.ShowAutoComplete "additem clear"
        SciSimple1.AddCallTip "additem(test)" & vbCrLf & "additem(test,test)"
    ElseIf className = "app" Or prevWord = "app" Then
        SciSimple1.ShowAutoComplete "getclipboard setclipboard askvalue openfiledialog savefiledialog exec list benchmark enableIDADebugMessages"
    ElseIf className = "xx" Then
        SciSimple1.ShowCallTip "test(xx,yy)" & vbCrLf & "test(yy)" & vbCrLf & "test(zz)"
    ElseIf className = "yy" Then
         multi.InitAndShow Array("test(xx,yy)", "test(yy)", "test(zz)", "test(kksssssssssssss)")
    End If
        
    
End Sub

Private Sub SciSimple1_OnError(Number As String, Description As String)
    MsgBox "SciSimple Error: " & Description
End Sub

Private Sub SciSimple1_MarginClick(lline As Long, Position As Long, margin As Long, modifiers As Long)
    SciSimple1.SetMarker lline
    Text1 = lline
End Sub

Private Sub SciSimple1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    SciSimple1.hilightClear
    Me.Caption = SciSimple1.hilightWord(SciSimple1.CurrentWord) & " word matches"
End Sub
