VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15840
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   15840
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider Slider1 
      Height          =   435
      Left            =   2220
      TabIndex        =   2
      Top             =   180
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   767
      _Version        =   393216
      Max             =   128
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14820
      Top             =   8160
   End
   Begin sci2.SciSimple sc 
      Height          =   8775
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   15478
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'used to make sense of the .bin and Highlighter type see notes.txt for
'results

Dim j As Long

Private Type Highlighter
  StyleBold(127) As Long
  StyleItalic(127) As Long
  StyleUnderline(127) As Long
  StyleVisible(127) As Long
  StyleEOLFilled(127) As Long
  StyleFore(127) As Long
  StyleBack(127) As Long
  StyleSize(127) As Long
  StyleFont(127) As String
  StyleName(127) As String
  Keywords(7) As String
  strFilter As String
  strComment As String
  strName As String
  iLang As Long
  strFile As String
End Type

Const SCLEX_CPP = 3
Const SCLEX_HTML = 4
Const SCLEX_XML = 5
Const SCLEX_SQL = 7
Const SCLEX_VB = 8
Const SCLEX_Asasm = 9
Const SCLEX_ASM = 34
Const SCLEX_CPPNOCASE = 35
Const SCLEX_PHPSCRIPT = 69
   
Enum vbIndexes
    SCE_B_DEFAULT = 0
    SCE_B_COMMENT = 1
    SCE_B_NUMBER = 2
    SCE_B_KEYWORD = 3
    SCE_B_STRING = 4
    SCE_B_PREPROCESSOR = 5
    SCE_B_OPERATOR = 6
    SCE_B_IDENTIFIER = 7
    SCE_B_DATE = 8
    SCE_B_STRINGEOL = 9
    SCE_B_KEYWORD2 = 10
    SCE_B_KEYWORD3 = 11
    SCE_B_KEYWORD4 = 12
    SCE_B_CONSTANT = 13
    SCE_B_ASM = 14
    SCE_B_LABEL = 15
    SCE_B_ERROR = 16
    SCE_B_HEXNUMBER = 17
    SCE_B_BINNUMBER = 18
    SCE_B_COMMENTBLOCK = 19
    SCE_B_DOCLINE = 20
    SCE_B_DOCBLOCK = 21
    SCE_B_DOCKEYWORD = 22
End Enum

Sub setVB(i As vbIndexes, _
    Optional fore As ColorConstants = vbBlack, _
    Optional back As ColorConstants = vbWhite, _
    Optional font As String = "Courier New", _
    Optional size As Long = 11, _
    Optional bold As Boolean = False, _
    Optional italic As Boolean = False, _
    Optional underline As Boolean = False, _
    Optional visible As Boolean = True)

    With sc.DirectSCI
           .StyleSetBold i, IIf(bold, 1, 0)
           .StyleSetItalic i, IIf(italic, 1, 0)
           .StyleSetUnderline i, IIf(underline, 1, 0)
           .StyleSetVisible i, IIf(visible, 1, 0)
           .StyleSetFont i, font
           .StyleSetFore i, fore
           .StyleSetBack i, back
           .StyleSetSize i, size
    End With
    
End Sub
Private Sub Form_Load()

  'sc.LoadFile App.Path & "\x.txt"
  'sc.DirectSCI.SetLexer SCLEX_CPP
  
  sc.LoadFile App.Path & "\vb.txt"
  sc.DirectSCI.SetLexer SCLEX_VB
  
  Dim i As Long, x As Long
  
  'On Error GoTo hell
    

  With sc.DirectSCI
     .ClearDocumentStyle
     .StyleSetBits 5
     .StyleClearAll
     
     If .GetLexer = SCLEX_CPP Then
        .SetKeyWords 0, "VCallHresult ImpAdCallFPR4 LateMemCall keywords0"
        .SetKeyWords 1, "NewIfNullPr keywords1"
        .SetKeyWords 2, "FStVarAd keywords2"
        .SetKeyWords 3, "keywords3"
        .SetKeyWords 4, "keywords4"
        .SetKeyWords 5, "keywords5"
        .SetKeyWords 6, "keywords6"
        .SetKeyWords 7, "keywords7"
    Else
        .SetKeyWords 0, "object begin beginproperty endproperty end" 'must be lowercase
        '.SetKeyWords 1, "begin"
        '.SetKeyWords 2, "beginproperty"
        '.SetKeyWords 3, "endproperty end"
        '.SetKeyWords 4, "end"
        
        setVB SCE_B_COMMENT, &H5500
        setVB SCE_B_CONSTANT, vbRed
        setVB SCE_B_HEXNUMBER, vbGreen
        setVB SCE_B_IDENTIFIER, &HA00000 'dark blue
        setVB SCE_B_KEYWORD, vbRed
        'setVB SCE_B_KEYWORD2, vbMagenta
        'setVB SCE_B_KEYWORD3, vbYellow
        'setVB SCE_B_KEYWORD4, vbCyan
        setVB SCE_B_STRING, &HC00090
        setVB SCE_B_LABEL, &HFF00FF 'magenta
    End If
    
     '.StyleSetFore 2, vbRed
     
     'scintinilla.iface
'     # Styles in range 32..38 are predefined for parts of the UI and are not used as normal styles.
'    # Style 39 is for future use.
'    enu StylesCommon = STYLE_
'    Val STYLE_DEFAULT = 32
'    Val STYLE_LINENUMBER = 33
'    Val STYLE_BRACELIGHT = 34
'    Val STYLE_BRACEBAD = 35
'    Val STYLE_CONTROLCHAR = 36
'    Val STYLE_INDENTGUIDE = 37
'    Val STYLE_CALLTIP = 38
'    Val STYLE_LASTPREDEFINED = 39
'    Val STYLE_MAX = 255
'
'     .StyleSetBack 32, Highlighters(x).StyleBack(32)                          'so these blocks here are nonsense?
'     .StyleSetFore 32, Highlighters(x).StyleFore(32)                          'STYLE_DEFAULT = 32
'     .StyleSetVisible 32, CLng(Highlighters(x).StyleVisible(32))              'but .StyleClearAll wipes out and for loop
'     .StyleSetEOLFilled 32, CLng(Highlighters(x).StyleEOLFilled(32))          'would override anyway
'     .StyleSetBold 32, CLng(Highlighters(x).StyleBold(32))
'     .StyleSetItalic 32, CLng(Highlighters(x).StyleItalic(32))
'     .StyleSetUnderline 32, CLng(Highlighters(x).StyleUnderline(32))
'     .StyleSetFont 32, Highlighters(x).StyleFont(32)
'     .StyleSetSize 32, Highlighters(x).StyleSize(32)
'                                                            <-- wipes out all before it..so why is this block here
'

'     .StyleSetFore 35, .misc.BraceBadFore
'     .StyleSetFore 34, .misc.BraceMatchFore
'     .StyleSetBack 35, .misc.BraceBadBack
'     .StyleSetBack 34, .misc.BraceMatchBack
'     .StyleSetBold 35, .misc.BraceMatchBold
'     .StyleSetBold 34, .misc.BraceMatchBold
'     .StyleSetItalic 35, .misc.BraceMatchItalic
'     .StyleSetItalic 34, .misc.BraceMatchItalic
'     .StyleSetUnderline 35, .misc.BraceMatchUnderline
'     .StyleSetUnderline 34, .misc.BraceMatchUnderline
     
     .Colourise 0, -1
     '.currentHighlighter = strHighlighter
  End With

'Timer1.Enabled = True

End Sub
 

Private Sub Slider1_Change()
        sc.DirectSCI.StyleSetFore j, vbBlack
        Label1.Caption = Slider1.value
        sc.DirectSCI.StyleSetFore Slider1.value, vbRed
        sc.DirectSCI.Colourise 0, -1
        j = Slider1.value
End Sub

'cpp lexer (ilang3)
'5=keywords0, 16 = keywords1, 33 line numbers in gutter, 7 strings, 4 most offsets and bytecode 2 = comments
'19 keywords3, 11 to much crap, 1 multiline comment, 6 double quoted string
'10 braces amd brackets, +, 14 regex

'scilexer.h why all diff lexer constants in one file?
'#define SCE_C_DEFAULT 0
'#define SCE_C_COMMENT 1
'#define SCE_C_COMMENTLINE 2
'#define SCE_C_COMMENTDOC 3
'#define SCE_C_NUMBER 4
'#define SCE_C_WORD 5
'#define SCE_C_STRING 6
'#define SCE_C_CHARACTER 7
'#define SCE_C_UUID 8
'#define SCE_C_PREPROCESSOR 9
'#define SCE_C_OPERATOR 10
'#define SCE_C_IDENTIFIER 11
'#define SCE_C_STRINGEOL 12
'#define SCE_C_VERBATIM 13
'#define SCE_C_REGEX 14
'#define SCE_C_COMMENTLINEDOC 15
'#define SCE_C_WORD2 16
'#define SCE_C_COMMENTDOCKEYWORD 17
'#define SCE_C_COMMENTDOCKEYWORDERROR 18
'#define SCE_C_GLOBALCLASS 19
'#define SCE_C_STRINGRAW 20
'#define SCE_C_TRIPLEVERBATIM 21
'#define SCE_C_HASHQUOTEDSTRING 22
'#define SCE_C_PREPROCESSORCOMMENT 23
'#define SCE_C_PREPROCESSORCOMMENTDOC 24
'

Private Sub Timer1_Timer()

    If j > 127 Then
        Timer1.Enabled = False
        Me.Caption = ""
    Else
        sc.DirectSCI.StyleSetFore j, vbBlack
        j = j + 1
        Label1.Caption = j
        sc.DirectSCI.StyleSetFore j, vbRed
        sc.DirectSCI.Colourise 0, -1
    End If
    
End Sub
