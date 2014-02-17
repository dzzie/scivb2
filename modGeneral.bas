Attribute VB_Name = "modGeneral"
Option Explicit

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
Public Const m_def_WordWrap = 1
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
Public Const m_def_FoldHi = 0
Public Const m_def_FoldLo = 0

Public Const m_def_AutoCompleteStart = "."
Public Const m_def_AutoCompleteOnCTRLSpace = True
Public Const m_def_AutoCompleteString = "if then else"
Public Const m_def_AutoShowAutoComplete = 0

Public Const m_def_BookmarkBack = vbBlack
Public Const m_def_BookMarkFore = vbWhite
Public Const m_def_MarkerBack = vbBlack
Public Const m_def_MarkerFore = vbWhite

Public Const m_def_Gutter0Type = 1
Public Const m_def_Gutter0Width = 20
Public Const m_def_Gutter1Type = 0
Public Const m_def_Gutter1Width = 24
Public Const m_def_Gutter2Type = 0
Public Const m_def_Gutter2Width = 13

Public Enum dcShiftDirection
    lLeft = -1
    lRight = 0
End Enum

Global Const LANG_US = &H409


Public Function FileExists(strFile As String) As Boolean
  ' This is a generic function that uses the dir command
  ' to return a boolean value (true/false) on if a file exists.
  If Dir(strFile) = "" Then
    FileExists = False
  Else
    FileExists = True
  End If
End Function

Public Function IsNumericKey(KeyAscii As Integer) As Integer
  IsNumericKey = KeyAscii
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) Then KeyAscii = 0
End Function

Public Function Shift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long, ByVal lDirectionToShift As dcShiftDirection) As Long

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

' This function is utilized to return the modified position of the
' mousecursor on a window
Public Function GetWindowCursorPos(Window As Long) As POINTAPI
  Dim lP As POINTAPI
  Dim rct As RECT
  GetCursorPos lP
  GetWindowRect Window, rct
  GetWindowCursorPos.x = lP.x - rct.Left
  If GetWindowCursorPos.x < 0 Then GetWindowCursorPos.x = 0
  GetWindowCursorPos.Y = lP.Y - rct.Top
  If GetWindowCursorPos.Y < 0 Then GetWindowCursorPos.Y = 0
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
    For i = 0 To UBound(bVal())
      Byte2Str = Byte2Str & Chr(bVal(i))
    Next i
  End If
End Function

Public Function ShellDocument(sDocName As String, _
                    Optional ByVal action As String = "Open", _
                    Optional ByVal Parameters As String = vbNullString, _
                    Optional ByVal Directory As String = vbNullString, _
                    Optional ByVal WindowState As StartWindowState) As Boolean
    Dim Response
    Response = ShellExecute(&O0, action, sDocName, Parameters, Directory, WindowState)
    Select Case Response
        Case Is < 33
            ShellDocument = False
        Case Else
            ShellDocument = True
    End Select
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