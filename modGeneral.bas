Attribute VB_Name = "modGeneral"
Option Explicit

'----------------------------------- [ default sci init values ] ----------------------------------------
Public Const m_def_LineNumbers = 1
Public Const m_def_TabWidth = 4
Public Const m_def_CaretForeColor = vbBlack
Public Const m_def_CaretWidth = 2
Public Const m_def_EOLMode = 0 'CRLF
Public Const m_def_CodePage = 0
Public Const m_def_ContextMenu = 1
Public Const m_def_IgnoreAutoCompleteCase = 1
Public Const m_def_ReadOnly = 0
Public Const m_def_ScrollWidth = 2000
Public Const m_def_ShowFlags = 1
Public Const m_def_Text = "0"
Public Const m_def_SelText = "0"
Public Const m_def_ClearUndoAfterSave = 1
Public Const m_def_EndAtLastLine = 0
Public Const m_def_OverType = 0
Public Const m_def_ScrollBarH = 1
Public Const m_def_ScrollBarV = 1
Public Const m_def_ViewEOL = 0
Public Const m_def_ViewWhiteSpace = 0
Public Const m_def_ShowCallTips = 1
Public Const m_def_EdgeColor = &HE0E0E0
Public Const m_def_EdgeColumn = 0
Public Const m_def_EdgeMode = 0
Public Const m_def_EOL = 0
Public Const m_def_UseTabs = 0
Public Const m_def_WordWrap = 1 '0=none, 1 = wrap to word, 2=wrap to char (unused)
Public Const m_def_MarginFore = vbBlack
Public Const m_def_MarginBack = &HE0E0E0
Public Const m_def_LineBackColor = vbYellow
Public Const m_def_LineVisible = 0

Public Const m_def_AutoCloseQuotes = 0
Public Const m_def_AutoCloseBraces = 0

Public Const m_def_BraceMatchBold = 1
Public Const m_def_BraceMatchItalic = 0
Public Const m_def_BraceMatchUnderline = 0
Public Const m_def_BraceMatchBack = vbWhite
Public Const m_def_BraceBadBack = vbWhite
Public Const m_def_BraceMatch = vbBlue
Public Const m_def_BraceBad = vbRed
Public Const m_def_BraceHighlight = 1
Public Const m_def_HighlightBraces = 1

Public Const m_def_SelStart = 0
Public Const m_def_SelEnd = 0
Public Const m_def_SelBack = &HFFC0C0
Public Const m_def_SelFore = vbBlack

Public Const m_def_IndentationGuide = 0
Public Const m_def_IndentWidth = 4
Public Const m_def_MaintainIndentation = 1
Public Const m_def_TabIndents = 1
Public Const m_def_BackSpaceUnIndents = 1

Public Const m_def_Folding = 1
Public Const m_def_FoldAtElse = 0
Public Const m_def_FoldMarker = 2
Public Const m_def_FoldComment = True
Public Const m_def_FoldCompact = False
Public Const m_def_FoldHTML = False
'Public Const m_def_FoldHi = 0
'Public Const m_def_FoldLo = 0

Public Const m_def_AutoCompleteStart = "."
Public Const m_def_AutoCompleteOnCTRLSpace = True
Public Const m_def_AutoCompleteString = "if then else"
Public Const m_def_AutoShowAutoComplete = 0

'Public Const m_def_BookmarkBack = vbBlack
'Public Const m_def_BookMarkFore = vbWhite
Public Const m_def_MarkerBack = vbBlack
Public Const m_def_MarkerFore = vbWhite

Public Const m_def_Gutter0Type = 1
Public Const m_def_Gutter0Width = 50
Public Const m_def_Gutter1Type = 0
Public Const m_def_Gutter1Width = 13
Public Const m_def_Gutter2Type = 0
Public Const m_def_Gutter2Width = 13


'----------------------------------- [ end default sci init values ] ----------------------------------------

Private Enum dcShiftDirection
    lLeft = -1
    lRight = 0
End Enum

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lP As Long) As Long
Public Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, lP As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lP As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal m As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal flags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function SendMessageStringString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal iMsg As Long, ByVal str1 As String, ByVal str1 As String) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const SC_CP_UTF8 = 65001
Public Const GWL_WNDPROC = (-4)
Public Const MK_RBUTTON = &H2
Public Const MK_LBUTTON = &H1
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_HOME = &H24
Public Const VK_DOWN = &H28
Public Const VK_END = &H23
Public Const VK_UP = &H26
Public Const WM_USER As Long = &H400
Public Const EM_FORMATRANGE As Long = WM_USER + 57
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_PRESSED As Integer = &H1000
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOOWNERZORDER As Long = &H200
Public Const SWP_NOCOPYBITS = &H100
Public Const MK_CONTROL = &H8
Public Const MK_SHIFT = &H4
Public Const VK_SHIFT = &H10&
Public Const VK_CONTROL = &H11&
Public Const VK_MENU = &H12& ' Alt key
Public Const CB_FINDSTRING = &H14C
Public Const ALL_MESSAGES = -1
Public Const ERROR_SUCCESS = 0&
Global Const LANG_US = &H409

'Public Type tagInitCommonControlsEx
'   lngSize As Long
'   lngICC As Long
'End Type
'
'Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
'Public Const ICC_USEREX_CLASSES = &H200



Public Function GetUpper(varArray As Variant) As Long
    Dim Upper As Long
    On Error Resume Next
    
    Upper = UBound(varArray)
    If Err.Number Then
         If Err.Number = 9 Then
              Upper = 0
         Else
              With Err
                   MsgBox "Error:" & .Number & "-" & .Description
              End With
              Exit Function
         End If
    Else
         Upper = UBound(varArray) + 1
    End If
    
    On Error GoTo 0
    GetUpper = Upper
End Function


Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function


Public Function FileSize(fPath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fPath)
    
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



Function GetParentFolder(path) As String
    On Error Resume Next
    Dim tmp() As String, ub
    If Len(path) = 0 Then Exit Function
    If InStr(path, "\") < 1 Then
        GetParentFolder = path
    Else
        tmp = Split(path, "\")
        ub = tmp(UBound(tmp))
        GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
    End If
End Function

Public Function GetLoadedSciLexerPath() As String
     Dim h As Long, ret As String
     ret = Space(500)
     h = GetModuleHandle("SciLexer.dll")
     h = GetModuleFileName(h, ret, 500)
     If h > 0 Then ret = Mid(ret, 1, h)
     GetLoadedSciLexerPath = ret
End Function

Public Function GetFileVersion(Optional ByVal PathWithFilename As String) As String
    ' return file-properties of given file  (EXE , DLL , OCX)
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then Exit Function
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim b() As Byte
    Dim b2() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strTemp As String
    Dim n As Long
    
    ReDim b2(500)
    
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen <= 0 Then Exit Function
    
    ReDim b(lngBufferlen)
    lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, b(0))
    If lngRc = 0 Then Exit Function
    
    lngRc = VerQueryValue(b(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    MoveMemory b2(0), lngVerPointer, lngBufferlen
    lngHexNumber = b2(2) + b2(3) * &H100 + b2(0) * &H10000 + b2(1) * &H1000000
    strLangCharset = Right("0000000" & Hex(lngHexNumber), 8)
    
    strBuffer = String$(800, 0)
    strTemp = "\StringFileInfo\" & strLangCharset & "\FileVersion"
    lngRc = VerQueryValue(b(0), strTemp, lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    lstrcpy strBuffer, lngVerPointer
    n = InStr(strBuffer, Chr(0)) - 1
    If n > 0 Then
        strBuffer = Mid$(strBuffer, 1, n)
        GetFileVersion = strBuffer
    End If
   
End Function


Public Function FileExists(strFile As String) As Boolean
  If Len(strFile) = 0 Then Exit Function
  If Dir(strFile) <> "" Then FileExists = True
End Function

Public Function IsNumericKey(KeyAscii As Integer) As Integer
  IsNumericKey = KeyAscii
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) Then KeyAscii = 0
End Function

Private Function Shift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long, ByVal lDirectionToShift As dcShiftDirection) As Long

    Const ksCallname As String = "Shift"
    On Error GoTo Procedure_Error
    Dim LShift As Long

    If lDirectionToShift Then 'shift left
        LShift = lValue * (2 ^ lNumberOfBitsToShift)
    Else 'shift right
        LShift = lValue \ (2 ^ lNumberOfBitsToShift)
    End If

    
Procedure_Exit:
    Shift = LShift
    Exit Function
    
Procedure_Error:
    Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Procedure_Exit
End Function

Public Function LShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long

    Const ksCallname As String = "LShift"
    On Error GoTo Procedure_Error
    LShift = Shift(lValue, lNumberOfBitsToShift, lLeft)
    
Procedure_Exit:
    Exit Function
    
Procedure_Error:
    Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Procedure_Exit
End Function

Public Function GET_X_LPARAM(ByVal lParam As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(lParam), 8)
    GET_X_LPARAM = CLng("&H" & Right(hexstr, 4))
End Function

Public Function GET_Y_LPARAM(ByVal lParam As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(lParam), 8)
    GET_Y_LPARAM = CLng("&H" & Left(hexstr, 4))
End Function


Function GetSHIFT() As Long

    'This function returns the state of the
    '     SHIFT, CONTROL and ALT keys
    'It does not distinguish the difference
    '     in left or right
    'Return value:
    'Bit 0=1 if pressed)
    Dim KS As Long
    Dim RetVal As Long
    KS = 0
    RetVal = GetKeyState(VK_SHIFT)


    If (RetVal And 32768) <> 0 Then
        KS = KS Or 1
    End If

    GetSHIFT = KS
End Function

Public Function piGetShiftState() As Integer
Dim iR As Integer
Dim lR As Long
Dim lKey As Long
    iR = iR Or (-1 * pbKeyIsPressed(VK_SHIFT))
    iR = iR Or (-2 * pbKeyIsPressed(VK_MENU))
    iR = iR Or (-4 * pbKeyIsPressed(VK_CONTROL))
    piGetShiftState = iR

End Function

Private Function pbKeyIsPressed(ByVal nVirtKeyCode As KeyCodeConstants) As Boolean
Dim lR As Long
    lR = GetAsyncKeyState(nVirtKeyCode)
    If (lR And &H8000&) = &H8000& Then
        pbKeyIsPressed = True
    End If
End Function

Private Sub pGetHiWordLoWord(ByVal lValue As Long, ByRef lHiWord As Long, ByRef lLoWord As Long)
    lHiWord = lValue \ &H10000
    lLoWord = (lValue And &HFFFF&)
End Sub

Public Function Max(a As Long, b As Long) As Long
  If a > b Then
    Max = a
  Else
    Max = b
  End If
End Function


Public Function Byte2Str(bVal() As Byte) As String
  Dim i As Long
  If GetUpper(bVal) <> 0 Then
'    For i = 0 To UBound(bVal())
'      Byte2Str = Byte2Str & Chr(bVal(i))
'    Next i
    Byte2Str = StrConv(bVal, vbUnicode, LANG_US)
  End If
End Function

'START_HIDDEN = 0, START_NORMAL = 4, START_MINIMIZED = 2, START_MAXIMIZED = 3
Public Function ShellDocument(sDocName As String, Optional ByVal action As String = "Open", Optional ByVal Parameters As String = vbNullString, Optional ByVal Directory As String = vbNullString, Optional ByVal WindowState As Long = 4) As Boolean
    If ShellExecute(&O0, action, sDocName, Parameters, Directory, WindowState) >= 33 Then ShellDocument = True
End Function

Sub SaveMySetting(key, Value)
    SaveSetting App.EXEName, "Settings", key, Value
End Sub

Function GetMySetting(key, Optional defaultval = "")
    GetMySetting = GetSetting(App.EXEName, "Settings", key, defaultval)
End Function

Sub FormPos(fform As Form, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.Name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.Name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Function isIDE() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIDE = False
    Exit Function
hell: isIDE = True
End Function

Function CountOccurancesOfChar(SearchText As String, SearchChar As String) As Integer

    Dim lCtr As Integer
    
    CountOccurancesOfChar = 0

    For lCtr = 1 To Len(SearchText)
        If StrComp(Mid(SearchText, lCtr, 1), SearchChar) = 0 Then
            CountOccurancesOfChar = CountOccurancesOfChar + 1
        End If
    Next

End Function

Function ReturnPositionOfOcurrance(SearchText As String, SearchChar As String, ByVal pPos As Integer) As Integer
    
    Dim lCtr As Integer
    ReturnPositionOfOcurrance = InStr(1, SearchText, "(") + 1

    If pPos <> 0 Then
        For lCtr = InStr(1, SearchText, "(") To Len(SearchText)
        If StrComp(Mid(SearchText, lCtr, 1), SearchChar) = 0 Then
                ReturnPositionOfOcurrance = lCtr
                pPos = pPos - 1
                If pPos = 0 Then
                    Exit Function
                End If
            End If
        Next

        ReturnPositionOfOcurrance = InStr(1, SearchText, ")") - 1

    End If
    
End Function

Function IsBrace(ch As Long) As Boolean
    IsBrace = (ch = 40 Or ch = 41 Or ch = 60 Or ch = 62 Or ch = 91 Or ch = 93 Or ch = 123 Or ch = 125)
End Function

Function MatchBrace(ch As String) As String
  If ch = "<" Then MatchBrace = ">"
  If ch = "(" Then MatchBrace = ")"
  If ch = "[" Then MatchBrace = "]"
  If ch = "{" Then MatchBrace = "}"
End Function

Function ReadFile(filename) As String 'this one should be binary safe...
  On Error GoTo hell
  Dim f As Long, b() As Byte
  f = FreeFile
  Open filename For Binary As #f
  ReDim b(LOF(f) - 1)
  Get f, , b()
  Close #f
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
  Exit Function
hell:   ReadFile = ""
End Function

Function writeFile(path As String, it As String) As Boolean   'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte, f As Long
    If FileExists(path) Then Kill path
    f = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #f
    Put f, , b()
    Close f
    writeFile = True
    Exit Function
hell: writeFile = False
End Function

Function GetBaseName(path) As String
    Dim tmp() As String, ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Sub Str2Byte(sInput As String, bOutput() As Byte)  '<--converted to strconv(lang_US) -dzzie
  ' This function is used to convert strings to bytes
  ' This comes in handy for saving the file.  It's also
  ' useful when dealing with certain things related to
  ' sending info to Scintilla

  'always Null terminated
  bOutput() = StrConv(sInput & Chr(0), vbFromUnicode, LANG_US)

'  Dim i As Long
'  ReDim bOutput(Len(sInput))
'
'  For i = 0 To Len(sInput) - 1
'    bOutput(i) = Asc(Mid(sInput, i + 1, 1))
'  Next i
'  bOutput(UBound(bOutput)) = 0 'Null terminated :)
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init: ReDim ary(0): ary(0) = Value
End Sub

Function RemoveTrailingNull(v As String) As String
    If Len(v) = 0 Then Exit Function
    If Right(v, 1) = Chr(0) Then
        If Len(v) = 1 Then
            RemoveTrailingNull = Empty
        Else
            RemoveTrailingNull = Left(v, Len(v) - 1)
        End If
    End If
End Function
