VERSION 5.00
Begin VB.UserControl SciSimple 
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   ForwardFocus    =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   7755
   ToolboxBitmap   =   "SciSimple.ctx":0000
   Begin VB.ListBox lstSort 
      Height          =   450
      Left            =   450
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "SciSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'NOTES:
'   could it be that the forward focus = false was the cause of my arrow key problems??
'   Available lexers compiled into _this_ scilexer dll:
'           asm vb vbscript sql cppnocase cpp hypertext xml phpscript

'Project no longer suffers from this IDE bug:
'  http://support.microsoft.com/kb/282233
'  BUG: Permission Denied Error Message When You Try to Recompile a Visual Basic
'         Project with a Public UDT and Binary Compatibility
'
' When you open a project and do anything that uses IntelliSense,
' you receive a "permission denied" error message when you try to recompile the project.
' The project defines a public User Defined Type (UDT) that it uses as a parameter to a
' public function, and binary compatibility is set.

Option Explicit

Implements iSubclass

Public misc As New CMiscSettings
Attribute misc.VB_VarDescription = "Misc class gives you access to some not often used features that would not be readily usable from DirectSCI access"
Public DirectSCI As New cDirectSCI
Attribute DirectSCI.VB_VarDescription = "Class that allows direct access to low level Scintilla API"

Event AutoCompleteEvent(className As String)
Event KeyPress(Char As Long)
Event DebugMsg(Msg As String)
Event KeyDown(KeyCode As Long, Shift As Long)
Event KeyUp(KeyCode As Long, Shift As Long)
Event key(ch As Long, modifiers As Long)
Event MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Event MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Event DoubleClick()
Event OnModified(Position As Long, modificationType As Long)
Event LineChanged(Position As Long)
Event UserListSelection(listType As Long, Text As String)   'Selected AutoComplete
Event CallTipClick(Position As Long)                        'Clicked a calltip
Event AutoCSelection(Text As String)                        'Auto Completed selected

Public Event MarginClick(lline As Long, Position As Long, margin As Long, modifiers As Long)
Public Event MouseDwellStart(lline As Long, Position As Long)
Public Event MouseDwellEnd(lline As Long, Position As Long)
Public Event NewLine()


Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


'=========[ scisimple private values ]====================
Private SCI As Long           ' hwnd for the Scintilla window
Private fWindowProc As Long   ' Proc Address of Scintilla.
Private SC As cSubclass       ' Subclass for Scintilla Messages

Public Enum SC_CODETYPE
  SC_DEFAULT = 0
  SC_CP_DBCS = 1
  SC_JAPANESE = 932
  SC_CHINESE = 936
  SC_KOREAN = 949
  SC_CP_UTF8 = 65001        ' Unicode support.
End Enum

Private hSciLexer As Long
Private m_hMod As Long
Private chStore As Long
Private mLastTopLine As Long

Private APIStrings() As String
Private ActiveCallTip As Integer 'no reason for this to be global?

' EOL Style Enum  (Scintilla supports Windows, Linux and Mac Line Endings)
Private Enum EOLStyle
  SC_EOL_CRLF = 0                     ' CR + LF
  SC_EOL_CR = 1                       ' CR
  sc_eol_lf = 2                       ' LF
End Enum

' Edge Style Enum (This is for a column edge)
Public Enum edge
  EdgeNone = 0
  EdgeLine = 1
  EdgeBackground = 2
End Enum

Public Enum FoldingStyle
  FoldMarkerArrow = 0
  foldMarkerBox = 1
  FoldMarkerCircle = 2
  FoldMarkerPlusMinus = 3
End Enum

Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    Code As Long
End Type

Private Type SCNotification
    NotifyHeader As NMHDR
    Position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    x As Long
    y As Long
End Type

Dim m_CodePage As SC_CODETYPE
Dim m_SelStart As Long
Dim m_SelEnd As Long
Dim m_IndentationGuide As Boolean
Dim m_FoldAtElse As Boolean
Dim m_FoldComment As Boolean
Dim m_FoldCompact As Boolean
Dim m_FoldHTML As Boolean
Dim m_AutoCompleteStart As String
Dim m_AutoCompleteOnCTRLSpace As Boolean
Dim m_AutoCompleteString As String
Dim m_AutoShowAutoComplete As Boolean
Dim m_ContextMenu As Boolean
Dim m_IgnoreAutoCompleteCase As Boolean
Dim m_LineNumbers As Boolean
Dim m_ReadOnly As Boolean
Dim m_ScrollWidth As Long
Dim m_ShowFlags As Boolean
Dim m_Text As String
Dim m_SelText As String
Dim m_MarginFore As OLE_COLOR
Dim m_MarginBack As OLE_COLOR
Dim m_FoldMarker As FoldingStyle
Dim m_AutoCloseBraces As Boolean
Dim m_AutoCloseQuotes As Boolean
Dim m_BraceHighlight As Boolean
Dim m_LineBackColor As OLE_COLOR
Dim m_LineVisible As Boolean
Dim m_CaretWidth As Long
Dim m_ClearUndoAfterSave As Boolean
Dim m_SelBack As OLE_COLOR
Dim m_SelFore As OLE_COLOR
Dim m_EndAtLastLine As Boolean
Dim m_OverType As Boolean
Dim m_ScrollBarH As Boolean
Dim m_ScrollBarV As Boolean
Dim m_ViewEOL As Boolean
Dim m_ViewWhiteSpace As Boolean
Dim m_ShowCallTips As Boolean
Dim m_EdgeColor As OLE_COLOR
Dim m_EdgeColumn As Long
Dim m_EdgeMode As edge
Dim m_EOL As EOLStyle
Dim m_Folding As Boolean
Dim m_MaintainIndentation As Boolean
Dim m_TabIndents As Boolean
Dim m_BackSpaceUnIndents As Boolean
Dim m_IndentWidth As Long
Dim m_UseTabs As Boolean
Dim m_WordWrap As Long '0 = none, 1 = wrap, 2 = wrap char? (unused)
Dim m_TabWidth As Long
Dim m_EOLMode As Long
Dim m_matchBraces
Dim m_CurrentHighlighter As String

'for Find/FindNext support
Private bRegEx As Boolean
Private bWholeWord As Boolean
Private bAutoSelectFinds As Boolean
Private bWordStart As Boolean
Private bCase As Boolean
Private strFind As String
Private bFindEvent As Boolean
Private LastFindPos As Long

Private bShowCallTips As Boolean
Private bShowFlags As Boolean
Private strAutoComplete As String
Private strAutoCompleteStart As String
Private bShowAutoComplete As Boolean
Private bRepLng As Boolean
Private bRepAll As Boolean
Private bReplaceFormActive As Boolean

Friend Property Let ReplaceFormActive(x As Boolean)
    bReplaceFormActive = x
End Property

 '=========================[ subclassing, initilization, and usercontrol stuff ]====================================

Property Get sciHWND() As Long
    sciHWND = SCI
End Property

Private Sub AttachHooks()
  Set SC = New cSubclass
  With SC

    '.Subclass UserControl.hwnd, Me
    '.AddMsg UserControl.hwnd, VK_LEFT, MSG_BEFORE

    .Subclass UserControl.hwnd, Me
    .AddMsg UserControl.hwnd, WM_NOTIFY, MSG_AFTER
    .AddMsg UserControl.hwnd, WM_SETFOCUS, MSG_AFTER
    .AddMsg UserControl.hwnd, WM_CLOSE, MSG_BEFORE
    .AddMsg UserControl.hwnd, WM_KEYDOWN, MSG_BEFORE '_AND_AFTER

    .Subclass SCI, Me
    .AddMsg SCI, WM_RBUTTONDOWN, MSG_AFTER
    .AddMsg SCI, WM_LBUTTONDOWN, MSG_AFTER
    .AddMsg SCI, WM_KEYDOWN, MSG_BEFORE '_AND_AFTER
    .AddMsg SCI, WM_KEYUP, MSG_AFTER
    .AddMsg SCI, WM_LBUTTONUP, MSG_AFTER
    .AddMsg SCI, WM_RBUTTONUP, MSG_AFTER
    .AddMsg SCI, WM_CHAR, MSG_BEFORE
    .AddMsg SCI, WM_COMMAND, MSG_BEFORE

  End With
End Sub

Private Sub HandleSciMsg(tHdr As NMHDR, scMsg As SCNotification)
    'Scintilla has given some information.
    'Let's see what it is and route it to the proper place.
    'Any commented with TODO have not been implimented yet.

    Dim strTmp As String
    Dim zPos As Long
    Dim chl As String, strMatch As String
    Dim lPos As Long
    Dim pos As Long, pos2 As Long
    
    Select Case tHdr.Code
            Case SCN_MODIFIED
                                RaiseEvent OnModified(scMsg.Position, scMsg.modificationType)
            'Case 2012
                                'RaiseEvent PosChanged(scMsg.Position)
            Case SCN_KEY
                                RaiseEvent key(scMsg.ch, scMsg.modifiers)
            Case SCN_STYLENEEDED
                                'RaiseEvent StyleNeeded(scMsg.Position)
            Case SCN_CHARADDED
                                'RaiseEvent CharAdded(scMsg.ch)
                                chStore = scMsg.ch
                                
                                If AutoCloseBraces Then
                                    chl = Chr(scMsg.ch)
                                    If chl = "(" Or chl = "[" Or chl = "{" Then
                                        strMatch = MatchBrace(chl)
                                        lPos = DirectSCI.GetCurPos
                                        DirectSCI.AddText 1, strMatch
                                        DirectSCI.SetSel lPos, lPos
                                    End If
                                End If
                                
                                If AutoCloseQuotes Then
                                    chl = Chr(scMsg.ch)
                                    If chl = """" Or chl = "'" Then
                                        If chl = """" Then
                                             strMatch = """"
                                        Else
                                             strMatch = "'"
                                        End If
                                        lPos = DirectSCI.GetCurPos
                                        DirectSCI.AddText 1, strMatch
                                        DirectSCI.SetSel lPos, lPos
                                    End If
                                End If
                                
                                'chl = scMsg.ch
                                If MaintainIndentation = True Then
                                    If scMsg.ch = 13 Or scMsg.ch = 10 Then
                                        MaintainIndent
                                    End If
                                End If
                                 
                                If bShowCallTips Then
                                     StartCallTip scMsg.ch
                                End If

            Case SCN_SAVEPOINTREACHED
                                'RaiseEvent SavePointReached
            Case SCN_SAVEPOINTLEFT
                                'RaiseEvent SavePointLeft
            Case SCN_MODIFYATTEMPTRO
              'TODO
            Case SCN_DOUBLECLICK
                                RaiseEvent DoubleClick
            Case SCN_UPDATEUI
                                
                                If m_BraceHighlight = False Then
                                    DirectSCI.BraceBadLight -1
                                    DirectSCI.BraceHighlight -1, -1
                                Else
                                    
                                    pos2 = INVALID_POSITION
                                    
                                    If IsBrace(DirectSCI.CharAtPos(DirectSCI.GetCurPos)) Then
                                        pos2 = DirectSCI.GetCurPos
                                    ElseIf IsBrace(DirectSCI.CharAtPos(DirectSCI.GetCurPos - 1)) Then
                                        pos2 = DirectSCI.GetCurPos - 1
                                    End If
                                    
                                    If pos2 <> INVALID_POSITION Then
                                        pos = SendMessage(SCI, SCI_BRACEMATCH, pos2, CLng(0))
                                        If pos = INVALID_POSITION Then
                                            Call SendEditor(SCI_BRACEBADLIGHT, pos2)
                                        Else
                                            Call SendEditor(SCI_BRACEHIGHLIGHT, pos, pos2)
                                            'If m_IndGuides Then
                                                Call SendEditor(SCI_SETHIGHLIGHTGUIDE, DirectSCI.GetColumn)
                                            'End If
                                        End If
                                    Else
                                        Call SendEditor(SCI_BRACEHIGHLIGHT, INVALID_POSITION, INVALID_POSITION)
                                    End If
                                    
                                End If
                                
                                If mLastTopLine <> Me.FirstVisibleLine Then
                                    mLastTopLine = Me.FirstVisibleLine
                                    RaiseEvent LineChanged(mLastTopLine)
                                End If
                                
                                'RaiseEvent UpdateUI
                                
                                
                                
            'Case SCN_MACRORECORD
                                '  HandleMacroCall scMsg.message, Chr(chStore)
                                '  RaiseEvent MacroRecord(scMsg.message, wParam)
                                
            Case SCN_MARGINCLICK
                                Dim lline As Long, lMargin As Long, lPosition As Long
                                lPosition = scMsg.Position
                                lline = SendEditor(SCI_LINEFROMPOSITION, lPosition)
                                lMargin = scMsg.margin
                                
                                If lMargin = MARGIN_SCRIPT_FOLD_INDEX Then
                                    Call SendEditor(SCI_TOGGLEFOLD, lline, 0)
                                End If
                                
                                'RaiseEvent MarginClick(scMsg.modifiers, scMsg.Position)
                                RaiseEvent MarginClick(lline, scMsg.Position, scMsg.margin, scMsg.modifiers)
                                
            Case SCN_NEEDSHOWN
                                'TODO
              
            Case SCN_CALLTIPCLICK
                                RaiseEvent CallTipClick(scMsg.Position)
                                
            Case SCN_PAINTED
                                'RaiseEvent Painted
                                
                                
            Case SCN_AUTOCSELECTION
                                strTmp = String(255, " ")
                                ConvCStringToVBString strTmp, scMsg.Text
                                zPos = InStr(strTmp, vbNullChar)
                                strTmp = Left(strTmp, zPos - 1)
                                RaiseEvent AutoCSelection(strTmp)
                                
            Case SCN_USERLISTSELECTION
                                strTmp = String(255, " ")
                                ConvCStringToVBString strTmp, scMsg.Text
                                zPos = InStr(strTmp, vbNullChar)
                                strTmp = Left(strTmp, zPos - 1)
                                RaiseEvent UserListSelection(scMsg.listType, strTmp)
                                
            Case SCN_DWELLSTART
                                lline = DirectSCI.SendEditor(SCI_LINEFROMPOSITION, scMsg.Position)
                                RaiseEvent MouseDwellStart(lline, scMsg.Position)
            Case SCN_DWELLEND
                                lline = DirectSCI.SendEditor(SCI_LINEFROMPOSITION, scMsg.Position)
                                RaiseEvent MouseDwellEnd(lline, scMsg.Position)

    End Select
    
End Sub

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
    On Error Resume Next
    
    Dim scMsg As SCNotification
    Dim tHdr As NMHDR
    Dim strTmp As String
    Dim Shift As Long
    Dim tmpStr As String
    Dim lP As POINTAPI
    Dim zPos As Long
    Dim chl As String, strMatch As String
    Dim lPos As Long
    Dim x As Long
        
    'this one is handled seperate so we can set breakpoints on the select and not see these..
    If uMsg = WM_NOTIFY Then
        CopyMemory scMsg, ByVal lParam, Len(scMsg)
        tHdr = scMsg.NotifyHeader
        If (tHdr.hwndFrom = SCI) Then HandleSciMsg tHdr, scMsg
        Exit Sub
    End If
                    
    Select Case uMsg

      Case WM_LBUTTONDOWN
                    RaiseEvent MouseDown(1, GetSHIFT(), GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
                    
      Case WM_CLOSE
                    'Detach ' Just to be safe detach it.
      Case WM_CHAR

                    If wParam = 32 And piGetShiftState = 4 Then 'CTRL Space
                        bHandled = True
                        lReturn = 0
                        strMatch = CurrentWord
                        'If Len(strMatch) > 0 Then RaiseEvent AutoCompleteEvent(strMatch)
                        RaiseEvent AutoCompleteEvent(strMatch) 'behavior changed 6.26.14 -dz
                    Else
                        bHandled = False
                        lReturn = 0
                        RaiseEvent KeyPress(wParam)
                    End If
                     
                                        
      Case WM_RBUTTONDOWN
                    lP = GetWindowCursorPos(SCI)
                    RaiseEvent MouseDown(2, GetSHIFT(), GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
                    
      Case WM_LBUTTONUP
                    If bReplaceFormActive And GetSHIFT() = 1 Then 'shift key pressed
                        If Me.SelLength > 0 Then frmReplace.SetFindText Me.SelText
                    End If
                    lP = GetWindowCursorPos(SCI)
                    RaiseEvent MouseUp(1, GetSHIFT(), GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
                    
      Case WM_RBUTTONUP
                    lP = GetWindowCursorPos(SCI)
                    RaiseEvent MouseUp(2, GetSHIFT(), GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
                    
      Case WM_KEYDOWN
                                           
                    If piGetShiftState = 4 Then 'CTRL Key
                    
                        If wParam = Asc("C") Or wParam = Asc("X") Then 'copy/cut
                            Clipboard.Clear
                            Clipboard.SetText Me.SelText
                            If wParam = Asc("X") Then SelText = ""
                        End If
                                                    
                        If wParam = Asc("V") Then SelText = Clipboard.GetText 'paste
                        If wParam = Asc("A") Then SelectAll
                        If wParam = Asc("Z") Then Undo
                        If wParam = Asc("Y") Then Redo
                        
                    End If
                    
                    If piGetShiftState = 5 Then
                        If wParam = 32 Then
                            StartCallTip Asc("(")
                        End If
                    End If
                    
                    If bShowCallTips And scMsg.ch <> 0 Then
                        StartCallTip scMsg.ch
                    End If
            
                    If wParam = 13 Then
                        RaiseEvent NewLine
                    End If
    
                    RaiseEvent KeyDown(wParam, piGetShiftState)
                    
      Case WM_KEYUP
                    
                    If wParam = 190 Then 'period
                        strMatch = CurrentWord
                        If Len(strMatch) > 0 Then RaiseEvent AutoCompleteEvent(strMatch)
                    End If
                    
                     If piGetShiftState = 4 Then 'CTRL Key
                        
                        If wParam = Asc("F") Or wParam = Asc("H") Then
                            frmReplace.LaunchReplaceForm Me
                        End If
                        
                        If Asc("G") = wParam Then
                            Call ShowGoto
                            bHandled = True
                            lReturn = 0
                            wParam = 0
                        End If
                        
                    End If
                    
                    If bShowCallTips Then
                        StartCallTip scMsg.ch
                    End If
                    
                    RaiseEvent KeyUp(wParam, piGetShiftState)
                    
      Case WM_SETFOCUS
                    DirectSCI.SetFocus
                    
    End Select


End Sub

'this is only called from initscintinilla right now..
Private Sub SetOptions()
        Dim i As Long

        DirectSCI.SetCaretFore m_def_CaretForeColor
        DirectSCI.SetCaretWidth m_def_CaretWidth
        
        DirectSCI.SetEdgeColour m_EdgeColor
        DirectSCI.SetEdgeColumn m_EdgeColumn
        DirectSCI.SetEdgeMode m_EdgeMode
        DirectSCI.SetIndentationGuides m_IndentationGuide
        DirectSCI.UsePopUp m_ContextMenu
        DirectSCI.SetReadOnly m_ReadOnly
        DirectSCI.SetEndAtLastLine m_EndAtLastLine
        DirectSCI.SetEOLMode m_EOL
        
        SendEditor SCI_SETCODEPAGE, m_CodePage, 0
        SetFoldMarker m_FoldMarker
        
        DirectSCI.SetMarginTypeN 0, misc.GutterType(gut0)
        DirectSCI.SetMarginTypeN 1, misc.GutterType(gut1)
        DirectSCI.SetMarginTypeN 2, misc.GutterType(gut2)
        
        If Folding = True Then
          DirectSCI.SetMarginWidthN 2, misc.GutterWidth(gut2)
        End If
        
        If LineNumbers = True Then
          DirectSCI.SetMarginWidthN 0, misc.GutterWidth(gut0)
        End If
        
        If ShowFlags = True Then
          DirectSCI.SetMarginWidthN 1, misc.GutterWidth(gut1)
        End If
        
        For i = 0 To 4
            DirectSCI.SetMarginSensitiveN i, 1
        Next
        DirectSCI.SetMouseDwellTime 600
        
        DirectSCI.SetCaretLineVisible m_LineVisible
        DirectSCI.SetCaretLineBack m_LineBackColor
        
        misc.MarkerBack = misc.MarkerBack
        misc.MarkerFore = misc.MarkerFore
        
        misc.BraceBadFore = misc.BraceBadFore
        misc.BraceMatchFore = misc.BraceMatchFore
        misc.BraceMatchBack = misc.BraceMatchBack
        misc.BraceBadBack = misc.BraceBadBack
        misc.BraceMatchBold = misc.BraceMatchBold
        misc.BraceMatchItalic = misc.BraceMatchItalic
        misc.BraceMatchUnderline = misc.BraceMatchUnderline
        
        'DirectSCI.SetMarkerBack 1, m_BookmarkBack
        'DirectSCI.SetMarkerFore 1, m_BookMarkFore
        
        DirectSCI.SetOvertype m_OverType
        DirectSCI.SetHScrollBar m_ScrollBarH
        DirectSCI.SetVScrollBar m_ScrollBarV
        DirectSCI.SetSelBack True, m_SelBack
        DirectSCI.SetSelFore True, m_SelFore
        DirectSCI.SetTabIndents m_TabIndents
        DirectSCI.SetUseTabs m_UseTabs
        DirectSCI.SetTabWidth m_IndentWidth
        DirectSCI.SetViewEOL m_ViewEOL
        DirectSCI.SetViewWS CLng(m_ViewWhiteSpace)
        DirectSCI.SetWrapMode m_WordWrap
        
        Folding = Folding
        ShowFlags = ShowFlags
        LineNumbers = LineNumbers
        InitFolding Folding
        
End Sub

Private Sub Detach()
  SC.UnSubAll
  Set SC = Nothing
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    MoveSCI 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub


Private Sub UserControl_Terminate()
    On Error GoTo Catch
    'Stop all subclassing
    'Detach
    'FreeLibrary m_hMod
    'FreeLibrary hSciLexer
Catch:
End Sub

Private Sub UserControl_Initialize()

        'On Error Resume Next
        
'        Dim iccex As tagInitCommonControlsEx 'I dont think this is required, only for themese support..
'        iccex.lngSize = LenB(iccex)
'        iccex.lngICC = ICC_USEREX_CLASSES
'        InitCommonControlsEx iccex
        
        'this is to prevent crash
        m_hMod = LoadLibrary("shell32.dll")
    
        misc.Initilize Me
        UserControl_InitProperties 'normally this would only be called when the usercontrol is first dropped on a form..
        InitScintilla
        
        Dim f As String
        f = App.path & "\java.hilighter"
        If FileExists(f) Then LoadHighlighter f

End Sub

Private Function InitScintilla() As Boolean
    On Error GoTo errHandler
    
    hSciLexer = LoadLibrary("SciLexer.DLL")
    If hSciLexer = 0 Then hSciLexer = LoadLibrary(App.path & "\SciLexer.DLL")
    If hSciLexer = 0 Then hSciLexer = LoadLibrary(App.path & "\..\SciLexer.DLL")
    
    If hSciLexer = 0 Then
      RaiseEvent DebugMsg("Failed to load SciLexer.DLL")
      Exit Function
    End If
    
    Set DirectSCI = New cDirectSCI
    SCI = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", "Scint.ocx", WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, UserControl.hwnd, 0, App.hInstance, 0)
    DirectSCI.SCI = SCI
    
    If SCI = 0 Then
      RaiseEvent DebugMsg("Failed to initilize Scintilla interface.")
      Exit Function
    End If
    
    fWindowProc = GetWindowLong(SCI, GWL_WNDPROC)
    AttachHooks
    DirectSCI.SetBackSpaceUnIndents BackSpaceUnIndents
    SetOptions
    RemoveHotKeys
    DirectSCI.SetPasteConvertEndings True
    DirectSCI.SetFocus
    InitScintilla = True
    
    Exit Function
errHandler:
    RaiseEvent DebugMsg("Error in InitScintilla: " & Err.Description)
End Function


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    bShowCallTips = m_def_ShowCallTips
    m_AutoCloseBraces = m_def_AutoCloseBraces
    m_TabWidth = m_def_TabWidth
    m_EOLMode = m_def_EOLMode
    m_EndAtLastLine = m_def_EndAtLastLine
    m_AutoCloseQuotes = m_def_AutoCloseQuotes
    m_LineBackColor = m_def_LineBackColor
    m_LineVisible = m_def_LineVisible
    m_ClearUndoAfterSave = m_def_ClearUndoAfterSave
    m_SelBack = m_def_SelBack
    m_SelFore = m_def_SelFore
    m_EndAtLastLine = m_def_EndAtLastLine
    m_OverType = m_def_OverType
    m_ScrollBarH = m_def_ScrollBarH
    m_ScrollBarV = m_def_ScrollBarV
    m_ViewEOL = m_def_ViewEOL
    m_ViewWhiteSpace = m_def_ViewWhiteSpace
    m_ShowCallTips = m_def_ShowCallTips
    m_EdgeColor = m_def_EdgeColor
    m_EdgeColumn = m_def_EdgeColumn
    m_EdgeMode = m_def_EdgeMode
    m_EOL = m_def_EOL
    m_Folding = m_def_Folding
    m_MaintainIndentation = m_def_MaintainIndentation
    m_TabIndents = m_def_TabIndents
    m_BackSpaceUnIndents = m_def_BackSpaceUnIndents
    m_IndentWidth = m_def_IndentWidth
    m_UseTabs = m_def_UseTabs
    m_WordWrap = m_def_WordWrap
    m_FoldMarker = m_def_FoldMarker
    m_MarginFore = m_def_MarginFore
    m_MarginBack = m_def_MarginBack
    m_Text = m_def_Text
    m_SelText = m_def_SelText
    m_AutoCompleteString = m_def_AutoCompleteString
    m_ContextMenu = m_def_ContextMenu
    m_IgnoreAutoCompleteCase = m_def_IgnoreAutoCompleteCase
    m_LineNumbers = m_def_LineNumbers
    m_ReadOnly = m_def_ReadOnly
    m_ScrollWidth = m_def_ScrollWidth
    m_ShowFlags = m_def_ShowFlags
    m_FoldAtElse = m_def_FoldAtElse
    m_FoldComment = m_def_FoldComment
    m_FoldCompact = m_def_FoldCompact
    m_FoldHTML = m_def_FoldHTML
    m_IndentationGuide = m_def_IndentationGuide
    m_SelStart = m_def_SelStart
    m_SelEnd = m_def_SelEnd
    m_BraceHighlight = m_def_BraceHighlight
    m_CodePage = m_def_CodePage
    
    With misc
        .BraceMatchFore = m_def_BraceMatch
        .BraceBadFore = m_def_BraceBad
        .BraceMatchBold = m_def_BraceMatchBold
        .BraceMatchItalic = m_def_BraceMatchItalic
        .BraceMatchUnderline = m_def_BraceMatchUnderline
        .BraceMatchBack = m_def_BraceMatchBack
        .BraceBadBack = m_def_BraceBadBack
        .GutterType(gut0) = m_def_Gutter0Type
        .GutterType(gut1) = m_def_Gutter1Type
        .GutterType(gut2) = m_def_Gutter2Type
        .GutterWidth(gut0) = m_def_Gutter0Width
        .GutterWidth(gut1) = m_def_Gutter1Width
        .GutterWidth(gut2) = m_def_Gutter2Width
        .MarkerBack = m_def_MarkerBack
        .MarkerFore = m_def_MarkerFore

    End With
    
End Sub

'======================================[ Hilighter code below ] =================================
Public Property Get currentHighlighter() As String
  currentHighlighter = m_CurrentHighlighter
End Property

'external users can not set this, for use from modHighlighter only..
Friend Property Let currentHighlighter(New_CurrentHighlighter As String)
  m_CurrentHighlighter = New_CurrentHighlighter
End Property

Public Function SetHighlighter(langName As String) As Boolean
  SetHighlighter = ModHighlighter.SetHighlighter(Me, langName)
End Function

Public Function LoadHighlighter(filePath As String, Optional andSetActive As Boolean = True) As Boolean
  On Error Resume Next
  Dim baseName As String
  baseName = GetBaseName(filePath)
 
  If Not ModHighlighter.LoadHighlighter(filePath) Then Exit Function
  If andSetActive Then
       If Not ModHighlighter.SetHighlighter(Me, baseName) Then Exit Function
  End If
  LoadHighlighter = True
End Function

Public Function HighlighterForExtension(fPath As String) As String
    HighlighterForExtension = ModHighlighter.HighlighterForExtension(fPath)
End Function

Public Function LoadHighlightersDir(dirPath As String) As Long
  On Error Resume Next
  LoadDirectory dirPath
  LoadHighlightersDir = ModHighlighter.HighLightersCount
End Function

Public Function ExportToHTML(filePath As String) As Boolean
    ExportToHTML = ExportToHTML3(filePath, Me)
End Function

Public Sub CommentBlock()
  CommentBlock2 Me
End Sub

Public Sub UncommentBlock()
  UncommentBlock2 Me
End Sub




'=======================================[ general functionality ]====================================================

Sub hilightClear()
    With DirectSCI
        .SendEditor SCI_INDICATORCLEARRANGE, 0, Len(Me.Text)
    End With
End Sub

Function hilightWord(sSearch As String, Optional color As Long = 0, Optional compare As VbCompareMethod = vbTextCompare) As Long
Attribute hilightWord.VB_Description = "Hilights all the instances of the search word. Does not work with older versions of SciLexer.dll"

    Dim lastIndex As Long
    Dim editorText As String
    Dim x As Long
    Dim hits As Long
    Dim curLine As Long
    
    Const Style = 7 'should i detect which version of scilexer is being used and change if old?
    Const inneralpha = 100
    Const borderalpha = 100
    
    lastIndex = 1
    x = 1
    
    If Len(sSearch) = 0 Then Exit Function
    
    LockWindowUpdate SCI
    editorText = Me.Text
    
    If color = 0 Then color = RGB(&HFF, &HFF, &H0)
    
    With DirectSCI
        .SendEditor SCI_SETINDICATORCURRENT, 9, 0
        .SendEditor SCI_INDICSETSTYLE, 9, Style
        .SendEditor SCI_INDICSETFORE, 9, color
        .SendEditor SCI_INDICSETALPHA, 9, inneralpha
        .SendEditor SCI_INDICSETOUTLINEALPHA, 9, borderalpha
        .SendEditor SCI_INDICSETUNDER, 9, vbBlack
        .SendEditor SCI_INDICATORCLEARRANGE, 0, Len(editorText)
    
        Do While x > 0
        
            x = InStr(lastIndex, editorText, sSearch, compare)
        
            If x + 2 = lastIndex Or x < 1 Or x >= Len(editorText) Then
                Exit Do
            Else
                lastIndex = x + Len(sSearch)
                DirectSCI.SendEditor SCI_INDICATORFILLRANGE, x - 1, Len(sSearch)
                hits = hits + 1
            End If
            
        Loop
        
        LockWindowUpdate 0
        hilightWord = hits
    End With
    
End Function

Property Get Version() As String
    Version = CompileVersionInfo(Me)
End Property

'auto close braces/quotes are handled by vb code in the subclass proc...
Public Property Get AutoCloseBraces() As Boolean    'When this is set to true braces <B>{, [, (</b> will be closed automatically.
    AutoCloseBraces = m_AutoCloseBraces
End Property

Public Property Let AutoCloseBraces(ByVal New_AutoCloseBraces As Boolean)
    m_AutoCloseBraces = New_AutoCloseBraces
    PropertyChanged "AutoCloseBraces"
End Property

Public Property Get AutoCloseQuotes() As Boolean    'When set to true quotes will automatically be closed.
    AutoCloseQuotes = m_AutoCloseQuotes
End Property

Public Property Let AutoCloseQuotes(ByVal New_AutoCloseQuotes As Boolean)
    m_AutoCloseQuotes = New_AutoCloseQuotes
    PropertyChanged "AutoCloseQuotes"
End Property

Sub GotoLineCentered(ByVal line As Long, Optional selected As Boolean = True)
    Dim mline As Long
    line = line - 1
    mline = line - CInt(DirectSCI.LinesOnScreen / 2)
    If mline > 0 Then FirstVisibleLine = mline
    GotoLine line
    If selected Then SelectLine
End Sub

Property Get FirstVisibleLine() As Long
    'returns the displayed line index, not absolute. if word wrap is on, it will be wrong..that was hard to find!
    
    Dim x As Long
    x = DirectSCI.GetFirstVisibleLine
    
    If Me.WordWrap Or Me.Folding Then
        x = DirectSCI.DocLineFromVisible(x)
    End If
    
    FirstVisibleLine = x
    
End Property

Property Let FirstVisibleLine(topLine As Long)

    GotoLine topLine + DirectSCI.LinesOnScreen + 5 'go past it
    GotoLine topLine   'now go to it and it will be topmost line..
    
End Property

Property Get VisibleLines() As Long
    VisibleLines = DirectSCI.LinesOnScreen
End Property

Property Get TotalLines() As Long
'    On Error Resume Next
'    Dim X As Long
'    X = UBound(Split(Me.Text, vbCrLf)) 'this does not handle vblf unix line endings...
'    If X = -1 Then X = 0
'    TotalLines = X
    TotalLines = DirectSCI.GetLineCount
End Property

Public Function FolderExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True
End Function

Public Function FileExists(strFile As String) As Boolean
  On Error Resume Next
  If Len(strFile) = 0 Then Exit Function
  If Dir(strFile) <> "" Then FileExists = True
End Function

Private Sub MoveSCI(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
     SetWindowPos SCI, 0, lLeft, lTop, lWidth / Screen.TwipsPerPixelX, lHeight / Screen.TwipsPerPixelY, SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Private Function SendEditor(ByVal Msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal lParam = 0) As Long
Attribute SendEditor.VB_Description = "sends a raw message to the scintilla editor"
    If VarType(lParam) = vbString Then
        SendEditor = SendMessageString(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), CStr(lParam))
    Else
        SendEditor = SendMessage(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), IIf(lParam = 0, CLng(lParam), lParam))
    End If
End Function

Public Property Get codePage() As SC_CODETYPE
    codePage = m_CodePage
End Property

Public Property Let codePage(ByVal New_CodePage As SC_CODETYPE)
    m_CodePage = New_CodePage
    PropertyChanged "CodePage"
    SendEditor SCI_SETCODEPAGE, New_CodePage, 0
End Property

Public Function GetLineText(ByVal lline As Long) As String
  'On Error Resume Next
  Dim txt As String
  Dim lLength As Long
  Dim i As Long
  Dim bByte() As Byte
  
  lLength = SendMessage(SCI, SCI_LINELENGTH, lline, 0)
  'lLength = lLength - 1 'By default this will tag on Chr(10) + chr(13) was failing on lines with only 1 char..
  
  If lLength > 0 Then
    ReDim bByte(0 To lLength)
    SendMessage SCI, SCI_GETLINE, lline, VarPtr(bByte(0))
    txt = Byte2Str(bByte())
    If Len(txt) > 1 Then If Right(txt, 1) = Chr(0) Then txt = Mid(txt, 1, Len(txt) - 1)
  Else
    txt = ""  'This line is 0 length
  End If
  
  GetLineText = txt
  
End Function

Public Property Get SelStart() As Long
    SelStart = DirectSCI.GetSelectionStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    m_SelStart = New_SelStart
    PropertyChanged "SelStart"
    DirectSCI.SetSelectionStart New_SelStart
End Property

Public Property Get SelEnd() As Long
    SelEnd = DirectSCI.GetSelectionEnd
End Property

Public Property Let SelEnd(ByVal New_SelEnd As Long)
    m_SelEnd = New_SelEnd
    PropertyChanged "SelEnd"
    DirectSCI.SetSelectionEnd New_SelEnd
End Property

Public Property Get SelLength() As Long
    On Error Resume Next
    SelLength = Len(SelText)
End Property

Public Property Let SelLength(vNewValue As Long)
    On Error Resume Next
    SelEnd = SelStart + vNewValue
End Property

Public Function GotoLine(line As Long) As Long
  DirectSCI.GotoLine line
End Function

Public Sub GotoLineColumn(iLine As Long, iCol As Long)
  Dim i As Long
  i = SendEditor(SCI_FINDCOLUMN, iLine, iCol)
  DirectSCI.SetSel i, i
End Sub

Public Function GotoCol(Column As Long) As Long
  GotoLineColumn CurrentLine, Column
End Function

Public Function SetFocus() As Long
  DirectSCI.SetFocus
End Function

Public Function Redo() As Long
  DirectSCI.Redo
End Function

Public Function Undo() As Long
  DirectSCI.Undo
End Function

Public Function Cut() As Long
  DirectSCI.Cut
End Function

Public Function Copy() As Long
  DirectSCI.Copy
End Function

Public Function Paste() As Long
  DirectSCI.Paste
End Function

Public Function SelectAll() As Long
  DirectSCI.SelectAll
End Function

Public Function SelectLine() As Long
  Dim curLine As Long
  curLine = CurrentLine
  DirectSCI.SetSel PositionFromLine(curLine), DirectSCI.GetLineEndPosition(curLine)
End Function


Public Property Get Text() As String    'Allows you to get and set the text of the scintilla window.
    Text = DirectSCI.GetText
End Property

Public Property Let Text(ByVal New_Text As String)
    Dim org_ro As Boolean
    If m_ReadOnly Then   'we have to let api access modify it even if its readonly from user
        org_ro = True
        ReadOnly = False
    End If
    m_Text = New_Text
    PropertyChanged "Text"
    DirectSCI.SetText New_Text
    If org_ro Then ReadOnly = True
    DirectSCI.SetFocus
End Property

Public Property Get SelText() As String 'Allows you to get and set the seltext of the scintilla window.
    SelText = DirectSCI.GetSelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Dim org_ro As Boolean
    If m_ReadOnly Then   'we have to let api access modify it even if its readonly from user
        org_ro = True
        ReadOnly = False
    End If
    m_SelText = New_SelText
    PropertyChanged "SelText"
    DirectSCI.SetSelText m_SelText
    If org_ro Then ReadOnly = True
    DirectSCI.SetFocus
End Property

Public Property Get ActiveLineBackColor() As OLE_COLOR    'Allows you to control the backcolor of the active line.
    ActiveLineBackColor = m_LineBackColor
End Property

Public Property Let ActiveLineBackColor(ByVal New_LineBackColor As OLE_COLOR)
    m_LineBackColor = New_LineBackColor
    PropertyChanged "LineBackColor"
    DirectSCI.SetCaretLineBack New_LineBackColor
End Property

Public Property Get HighLightActiveLine() As Boolean    'When set to true the active line will be highlighted using the color selected from LineBackColor.
   HighLightActiveLine = m_LineVisible
End Property

Public Property Let HighLightActiveLine(ByVal New_LineVisible As Boolean)
    m_LineVisible = New_LineVisible
    PropertyChanged "LineVisible"
    DirectSCI.SetCaretLineVisible m_LineVisible
End Property

Public Property Get MaintainIndentation() As Boolean 'If this is set to true the editor will automatically keep the previous line's indentation.
    MaintainIndentation = m_MaintainIndentation
End Property

Public Property Let MaintainIndentation(ByVal New_MaintainIndentation As Boolean)
    m_MaintainIndentation = New_MaintainIndentation
    PropertyChanged "MaintainIndentation"
End Property

Public Property Get ShowIndentationGuide() As Boolean   'If true indention guide's will be displayed.
    ShowIndentationGuide = m_IndentationGuide
End Property

Public Property Let ShowIndentationGuide(ByVal New_IndentationGuide As Boolean)
    m_IndentationGuide = New_IndentationGuide
    PropertyChanged "IndentationGuide"
    DirectSCI.SetIndentationGuides m_IndentationGuide
End Property

Private Function GetLineIndentPosition(lline As Long) As Long
  GetLineIndentPosition = SendEditor(SCI_GETLINEINDENTPOSITION, lline)
End Function

Public Property Get useTabs() As Boolean
    useTabs = m_UseTabs
End Property

Public Property Let useTabs(ByVal New_UseTabs As Boolean)
    m_UseTabs = New_UseTabs
    PropertyChanged "UseTabs"
    DirectSCI.SetUseTabs m_UseTabs
End Property

Public Property Get UseTabIndents() As Boolean 'If this is true tab inserts indent characters.  If it is set to false tab will insert spaces.
    UseTabIndents = m_TabIndents
End Property

Public Property Let UseTabIndents(ByVal New_TabIndents As Boolean)
    m_TabIndents = New_TabIndents
    PropertyChanged "TabIndents"
    DirectSCI.SetTabIndents m_TabIndents
End Property

Public Property Get BackSpaceUnIndents() As Boolean 'If tabindents is set to false, and BackSpaceUnIndents is set to true then the backspaceunindents will remove the same number of spaces as tab inserts.  If it's set to false then it will work normally.
    BackSpaceUnIndents = m_BackSpaceUnIndents
End Property

Public Property Let BackSpaceUnIndents(ByVal New_BackSpaceUnIndents As Boolean)
    m_BackSpaceUnIndents = New_BackSpaceUnIndents
    PropertyChanged "BackSpaceUnIndents"
    DirectSCI.SetBackSpaceUnIndents m_BackSpaceUnIndents
End Property

Public Property Get IndentWidth() As Long   'This controls the number of spaces Tab will indent.  IndentWidth only applies if <B>TabIndents</b> is set to false.
    IndentWidth = m_IndentWidth
End Property

Public Property Let IndentWidth(ByVal New_IndentWidth As Long)
    m_IndentWidth = New_IndentWidth
    PropertyChanged "IndentWidth"
    DirectSCI.SetTabWidth IndentWidth
    'SetIndent m_IndentWidth
End Property

Public Property Get AutoCompleteString() As String  'This store's the list which autocomplete will use.  Each word needs to be seperated by a space.
    AutoCompleteString = m_AutoCompleteString
End Property

Public Property Let AutoCompleteString(ByVal New_AutoCompleteString As String)
    m_AutoCompleteString = New_AutoCompleteString
    PropertyChanged "AutoCompleteString"
End Property

Public Property Get ContextMenu() As Boolean    'If set to true then the default Scintilla context menu will be displayed when a user right clicks on the window.  If this is set to false then no context menu will be displayed.  If you are utilizing a customer context menu then this should be set to false.
Attribute ContextMenu.VB_Description = "Use the default context menu or not. "
    ContextMenu = m_ContextMenu
End Property

Public Property Let ContextMenu(ByVal New_ContextMenu As Boolean)
    m_ContextMenu = New_ContextMenu
    PropertyChanged "ContextMenu"
    DirectSCI.UsePopUp m_ContextMenu
End Property
 
Private Function ConvertEOLMode()
  SendEditor SCI_CONVERTEOLS, DirectSCI.GetEOLMode
End Function

Public Sub ClearUndoBuffer()
  SendEditor SCI_EMPTYUNDOBUFFER
End Sub

Public Property Get LineNumbers() As Boolean    'If this is set to true then the first gutter will be visible and display line numbers.  If this is false then the first gutter will remain hidden.
    LineNumbers = m_LineNumbers
End Property

Public Property Let LineNumbers(ByVal New_LineNumbers As Boolean)
    m_LineNumbers = New_LineNumbers
    PropertyChanged "LineNumbers"
    If m_LineNumbers Then
      DirectSCI.SetMarginWidthN 0, misc.GutterWidth(gut0)
    Else
      DirectSCI.SetMarginWidthN 0, 0
    End If
End Property

Public Property Get ReadOnly() As Boolean  'This property allows you to set the readonly status of Scintilla.  When in readonly you can scroll the document, but no editing can be done.
    ReadOnly = m_ReadOnly
End Property

Public Property Let ReadOnly(ByVal New_ReadOnly As Boolean)
    m_ReadOnly = New_ReadOnly
    PropertyChanged "ReadOnly"
    DirectSCI.SetReadOnly m_ReadOnly
End Property

Public Property Get isDirty() As Boolean   'This is a read only property.  It allows you to get the modified status of the Scintilla window.
    isDirty = DirectSCI.GetModify
End Property

Public Property Get ShowFlags() As Boolean  'If this is true the second gutter will be displayed and Flags/Bookmarks will be displayed.
Attribute ShowFlags.VB_Description = "Enabled/Disables the flags gutter"
    ShowFlags = m_ShowFlags
End Property

Public Property Let ShowFlags(ByVal New_ShowFlags As Boolean)
    m_ShowFlags = New_ShowFlags
    PropertyChanged "ShowFlags"
    If m_ShowFlags Then
      DirectSCI.SetMarginWidthN 1, misc.GutterWidth(gut1)
    Else
      DirectSCI.SetMarginWidthN 1, 0
    End If
End Property

Public Property Get WordWrap() As Boolean 'If set to true the document will wrap lines which are longer than itself.  If false then it will dsiplay normally.
    WordWrap = IIf(m_WordWrap = 0, False, True)
End Property

Public Property Let WordWrap(ByVal wrap As Boolean)
    m_WordWrap = IIf(wrap, 1, 0)
    PropertyChanged "WordWrap"
    DirectSCI.SetWrapMode m_WordWrap
End Property

Public Property Get SelBack() As OLE_COLOR  'This allow's you to set the backcolor for selected text.
    SelBack = m_SelBack
End Property

Public Property Let SelBack(ByVal New_SelBack As OLE_COLOR)
    m_SelBack = New_SelBack
    PropertyChanged "SelBack"
    DirectSCI.SetSelBack True, m_SelBack
End Property

Public Property Get SelFore() As OLE_COLOR  'The allows you to control the fore color of the selected color.
    SelFore = m_SelFore
End Property

Public Property Let SelFore(ByVal New_SelFore As OLE_COLOR)
    m_SelFore = New_SelFore
    PropertyChanged "SelFore"
    DirectSCI.SetSelFore True, m_SelFore
End Property

Public Function PositionFromLine(lline As Long) As Long
  PositionFromLine = SendEditor(SCI_POSITIONFROMLINE, lline)
End Function

Public Sub SetCurrentPosition(lval As Long)
  SendEditor SCI_SETCURRENTPOS, lval
End Sub

Public Function CurrentLine() As Long
Attribute CurrentLine.VB_Description = "Gets the current line index"
  CurrentLine = DirectSCI.LineFromPosition(DirectSCI.GetCurPos)
End Function

Public Function GetCaretInLine() As Long
Attribute GetCaretInLine.VB_Description = "Gets the carret offset relative to the current line (different from GetColumn?)"
  Dim caret As Long, lineStart As Long, line As Long
  caret = DirectSCI.GetCurPos
  line = CurrentLine
  lineStart = PositionFromLine(line)
  GetCaretInLine = caret - lineStart
End Function

'takes a space delimited list of words and returns them alpha sorted
'sci editor requires the strings to be case _sensitive_ sorted
Private Function SortString(str As String) As String
  Dim x, tmp() As String
  On Error Resume Next
  tmp = Split(str, " ")
  If Not AryIsEmpty(tmp) Then
        lstSort.Clear
        For Each x In tmp 'list.sorted=true so it will auto sort the list for us :)
            If Len(x) > 0 Then lstSort.AddItem x
        Next
        Erase tmp
        For x = 0 To lstSort.ListCount()
            push tmp, lstSort.List(x)
        Next
        SortString = Trim(Join(tmp, " "))
  End If
End Function


Public Sub ShowAutoComplete(strVal As String)
  Dim i As Long
  
  If CanAutoCompleteCurWord(strVal) Then Exit Sub
  
  i = ToLastSpaceCount 'if last char was a " this will return 1 and autoc wont show...
  SendMessageString SCI, SCI_AUTOCSHOW, i, SortString(strVal)

End Sub

'if they hit ctrl-space in the middle of a word and there is only one
'match then this will autocomplete it and case correct it. also if the
'cursor is at the end of the word, and it is the only partial match for
'that first xx characters, then it will auto complete it (same behavior as vb6)
'this is wrapped on its own for error handling..
Private Function CanAutoCompleteCurWord(strVal As String) As Boolean

  On Error GoTo hell
    
  Dim iStart As Long, iEnd As Long, hits As Long
  Dim w As String, words() As String, word, matches() As String
  Dim lineStart As Long, firstCharIsDot As Boolean
  Const SCI_AUTOCCANCEL = 2101
  
  w = CurrentWordInternal(iStart, iEnd, firstCharIsDot)
  
  
  words = Split(strVal, " ")
  For Each word In words
        If Len(word) > 0 Then
            If LCase(word) = LCase(w) Or LCase(VBA.Left(word, Len(w))) = LCase(w) Then
                  push matches, word
            End If
        End If
  Next
  
  If Not AryIsEmpty(matches) Then
      If UBound(matches) = 0 Then 'only one match so we will autocomplete it and case correct
         lineStart = Me.PositionFromLine(CurrentLine)
         Me.SelStart = lineStart + iStart - 1
         Me.SelEnd = lineStart + iEnd
         Me.SelText = matches(0) & IIf(firstCharIsDot, ".", "")
         Me.SelLength = 0
         CanAutoCompleteCurWord = True
         SendMessage SCI, SCI_AUTOCCANCEL, 0, 0 'hide the auto select list if already shown
      End If                                    '(it was visible, then they hit ctrl space to complete cur selection)
  End If
    
hell:
    
End Function

Public Function CurrentWord() As String
    CurrentWord = CurrentWordInternal()
End Function

Private Function CurrentWordInternal(Optional iStart As Long, Optional iEnd As Long, Optional firstCharIsDot As Boolean) As String
    Dim line As String, x As Integer
    Dim newstr As String ', iPos As Integer, iStart As Long, iEnd As Long
    Dim i As Integer
    Dim c As String
    'Dim firstCharIsDot As Boolean
    
    line = GetLineText(CurrentLine())
    x = GetCaretInLine
    newstr = ""
    
    'parse the current line starting at the current cursor position and walking backwards..
    For i = x To 1 Step -1
        c = Mid(line, i, 1)
        If c = "." And i = x Then
            'ignore the class member access marker
            firstCharIsDot = True
        ElseIf InStr(1, CallTipWordCharacters, c) > 0 Then
            newstr = c & newstr
        Else
            If Asc(c) >= 32 Then   ' not valid character (and not whitespace)
                Exit For
            End If
        End If
    Next
    
    iStart = i + 1
    
    If firstCharIsDot Then
        iEnd = x
    Else
        'maybe they clicked in the middle of a word..now scan forward to find its end.
        For i = x + 1 To Len(line)
            c = Mid(line, i, 1)
            If InStr(1, CallTipWordCharacters, c) > 0 Then
                newstr = newstr & c
            Else
                Exit For
            End If
        Next
        
        iEnd = i - 1
    End If
    
    CurrentWordInternal = newstr

End Function

Public Function PreviousWord() As String
    Dim line As String, x As Integer
    Dim newstr As String
    Dim i As Integer
    Dim c As String
    Dim curWord As String
    Dim iStart As Long, iEnd As Long
    
    line = GetLineText(CurrentLine())
    x = GetCaretInLine
    newstr = ""
    
    'make sure to handle case if cursor is in middle of word, not just at end of a word..
    curWord = CurrentWordInternal(iStart, iEnd)
    'X = X - Len(curWord)
    x = iStart - 1
    
    'parse the current line starting at the current cursor position and walking backwards..
    For i = x To 1 Step -1
        c = Mid(line, i, 1)
        If c = "." And i = x Then
            'ignore the class member access marker
        ElseIf InStr(1, CallTipWordCharacters, c) > 0 Then
            newstr = c & newstr
        Else
            If Asc(c) >= 32 Then   ' not valid character (and not whitespace)
                Exit For
            End If
        End If
    Next

    PreviousWord = newstr

End Function

Public Function SaveFile(strFile As String) As Boolean
  On Error GoTo hell
  Dim str As String
  ConvertEOLMode
  str = DirectSCI.GetText
  writeFile strFile, str
  DirectSCI.SetSavePoint ' Remove the modified flag from scintilla
  If m_ClearUndoAfterSave Then ClearUndoBuffer
  SaveFile = True
  Exit Function
hell: SaveFile = False
End Function

Public Function LoadFile(strFile As String) As Boolean
  Dim str As String
  On Error GoTo hell
  
  If Not FileExists(strFile) Then Exit Function
  
  str = ReadFile(strFile)
  DirectSCI.SetText str
  ClearUndoBuffer
  DirectSCI.ConvertEOLs DirectSCI.GetEOLMode
  DirectSCI.SetFocus
  DirectSCI.GotoPos 0
  DirectSCI.SetSavePoint
  LoadFile = True
  
  Exit Function
hell: LoadFile = False
End Function



'===========================[ call tips ] ===================================

Public Function LoadCallTips(strFile As String) As Long
  On Error Resume Next
  Erase APIStrings  'Clear the old array
  If Not FileExists(strFile) Then Exit Function
  
  Dim tmp() As String, x
  
  tmp = Split(ReadFile(strFile), vbCrLf)
  For Each x In tmp
        x = Trim(x)
        If Len(x) > 0 And Left(x, 1) <> "#" And Left(x, 1) <> "'" Then
            push APIStrings, x
        End If
  Next
  
  LoadCallTips = UBound(APIStrings)
  
End Function

Public Function AddCallTip(functionPrototype As String)
    push APIStrings(), functionPrototype
End Function

Public Property Get DisplayCallTips() As Boolean   'If this is set to true then calltips will be displayed.  To use this you must also use <B>LoadAPIFile</b> to load an external API file which contains simple instructions to the editor on what calltips to display.
    DisplayCallTips = m_ShowCallTips
End Property

Public Property Let DisplayCallTips(ByVal New_ShowCallTips As Boolean)
    m_ShowCallTips = New_ShowCallTips
    PropertyChanged "ShowCallTips"
    bShowCallTips = m_ShowCallTips
End Property

Private Sub SetCallTipHighlight(lStart As Long, lEnd As Long)
  SendEditor SCI_CALLTIPSETHLT, lStart, lEnd
  Debug.Print lStart & " " & lEnd
  If lEnd = 0 Then Stop
End Sub

Public Sub StopCallTip()
  SendEditor SCI_CALLTIPCANCEL
End Sub

Public Sub ShowCallTip(strVal As String)
  Dim bByte() As Byte
  Str2Byte strVal, bByte
  Call SendEditor(SCI_CALLTIPSHOW, DirectSCI.GetCurPos, VarPtr(bByte(0)))
End Sub

Private Sub StartCallTip(ch As Long)
    ' This entire function is a bit of a hack.  It seems to work but it's very
    ' messy.  If anyone cleans it up please send me a new version so I can add
    ' it to this release.  Thanks :)
    Dim line As String, PartLine As String, i As Integer, x As Integer
    Dim newstr As String, iPos As Integer, iStart As Long, iEnd As Long
    Dim a, i2 As Integer
    
    If AryIsEmpty(APIStrings) Then Exit Sub
    
    If ch = Asc("(") Then
            line = GetLineText(CurrentLine())
            x = GetCaretInLine
            ' For those compilers that allow whitespace between function and parenthesis
            ' ignore whitespace
            For i2 = x - 1 To 1 Step -1
                If Asc(Mid(line, i2, 1)) < 33 And newstr <> "" Then    ' ignore whitespace before (
                    Exit For
                Else
                    If InStr(1, CallTipWordCharacters, Mid(line, i2, 1)) > 0 Then
                        newstr = Mid(line, i2, 1) & newstr
                    Else
                        If Asc(Mid(line, i2, 1)) > 33 Then   ' not valid character (and not whitespace)
                            Exit For
                        End If
                    End If
                End If
            Next i2
        
            If Len(newstr) = 0 Then   ' blank line ?
                StopCallTip
                Exit Sub
            End If
        
            newstr = newstr & "("    ' make it into a function name so no partial searches of other API functions
        
          ' Lookup the Function name in the API list
            For i = 0 To UBound(APIStrings)
                  If Left(LCase$(APIStrings(i)), Len(newstr)) = LCase$(newstr) Then ' case insensitive string
                          ActiveCallTip = i
              
                          If InStr(APIStrings(i), vbCrLf) > 0 Then 'multiline tip show whole thing - no cur arg highlights 12.12.19
                                ShowCallTip Left$(APIStrings(i), Len(APIStrings(i)))
                                'todo: highlight cur arg for each line..dont care for now rare for overloads and can use scroll
                                Exit Sub
                          Else
                                iPos = InStr(1, APIStrings(i), ")")
                                ShowCallTip Left$(APIStrings(i), iPos)  ' to end of function
                    
                                iPos = InStr(1, APIStrings(i), ",")
                                If iPos > 0 Then
                                    iStart = Len(newstr)
                                    iEnd = iPos - 1
                                    SetCallTipHighlight iStart, iEnd
                                    Exit For
                                Else
                                    ' single parameter ?
                                    If Len(newstr) + 1 <> Len(APIStrings(i)) Then
                                        iStart = Len(newstr)
                                        iEnd = Len(APIStrings(i)) - 1
                                        SetCallTipHighlight iStart, iEnd
                                        Exit For
                                    End If
                                End If
                         End If
                  End If
             Next
             
             Exit Sub
    End If
    
    ' Do we have a tip already active ?
    If DirectSCI.CallTipActive Then
            If ch = Asc(")") Then
                StopCallTip
            Else
                ' are we still in the current tooltip ?
                line = GetLineText(CurrentLine())
                x = GetCaretInLine
                iPos = InStrRev(line, "(", x)
                PartLine = Mid(line, iPos + 1, x - iPos) 'Get the chunk of the string were in
        
                If InStr(APIStrings(ActiveCallTip), vbCrLf) > 0 Then Exit Sub 'multiline exception for arg highlighting
                    
                If InStr(1, APIStrings(ActiveCallTip), ",") = 0 Then   ' only one param
                    iStart = InStr(1, APIStrings(ActiveCallTip), "(") - 1
                    iEnd = InStr(1, APIStrings(ActiveCallTip), ")") - 1
                Else
                    'Count which param
                    iPos = CountOccurancesOfChar(PartLine, ",")
                    'Highlight Param in calltip
                    iStart = ReturnPositionOfOcurrance(APIStrings(ActiveCallTip), ",", iPos) - 1
                    iEnd = ReturnPositionOfOcurrance(APIStrings(ActiveCallTip), ",", iPos + 1)
                End If
                SetCallTipHighlight iStart, iEnd
          End If
    End If
    
End Sub

'===========================[ end call tips ] ===================================







'===============================[ folding ] =======================================


Public Property Get Folding() As Boolean    'If true folding will be automatically handled.
    Folding = m_Folding
End Property

Public Property Let Folding(ByVal New_Folding As Boolean)
    m_Folding = New_Folding
    PropertyChanged "Folding"
    If m_Folding Then
      DirectSCI.SetMarginWidthN 2, misc.GutterWidth(gut2)
    Else
      DirectSCI.SetMarginWidthN 2, 0
    End If
    InitFolding New_Folding
End Property

Public Property Get FoldMarker() As FoldingStyle
    FoldMarker = m_FoldMarker
End Property

Public Property Let FoldMarker(ByVal New_FoldMarker As FoldingStyle)
    m_FoldMarker = New_FoldMarker
    PropertyChanged "FoldMarker"
    SetFoldMarker New_FoldMarker
End Property

Private Sub SetFoldMarker(Value As FoldingStyle)
    Select Case Value
    Case 1
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER)
    Case 2
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_CIRCLEMINUS)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_CIRCLEPLUS)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNERCURVE)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_CIRCLEPLUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_CIRCLEMINUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNERCURVE)
    Case 3
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_MINUS)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_PLUS)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_EMPTY)
    Case 0
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_ARROWDOWN)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_ARROW)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_EMPTY)
  End Select
End Sub

Private Sub InitFolding(EnableIt As Boolean)
  If EnableIt = True Then
    DirectSCI.SetProperty "fold", "1"
    DirectSCI.SetProperty "fold.compact", IIf(m_FoldCompact, "1", "0")
    DirectSCI.SetProperty "fold.comment", IIf(m_FoldComment, "1", "0")
    DirectSCI.SetProperty "fold.html", IIf(m_FoldHTML, "1", "0")
    If FoldAtElse = True Then
      DirectSCI.SetProperty "fold.at.else", "1"
    Else
      DirectSCI.SetProperty "fold.at.else", "0"
    End If
    'SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
    Call SendEditor(SCI_SETMARGINTYPEN, MARGIN_SCRIPT_FOLD_INDEX, SC_MARGIN_SYMBOL)
    Call SendEditor(SCI_SETMARGINMASKN, MARGIN_SCRIPT_FOLD_INDEX, SC_MASK_FOLDERS)
    'SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 20
    Call SendEditor(SCI_SETMARGINSENSITIVEN, MARGIN_SCRIPT_FOLD_INDEX, 1)
    Call SendEditor(SCI_SETFOLDFLAGS, 16, 0)
  Else
    DirectSCI.SetProperty "fold", "0"
    DirectSCI.SetProperty "fold.compact", 0
    DirectSCI.SetProperty "fold.html", "0"
    DirectSCI.SetProperty "fold.comment", "0"
    SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
    Call SendEditor(SCI_SETMARGINSENSITIVEN, MARGIN_SCRIPT_FOLD_INDEX, 0)
  End If
End Sub


Public Property Get FoldAtElse() As Boolean
    FoldAtElse = m_FoldAtElse
End Property

Public Property Let FoldAtElse(ByVal New_FoldAtElse As Boolean)
    m_FoldAtElse = New_FoldAtElse
    PropertyChanged "FoldAtElse"
    If FoldAtElse = True Then
      DirectSCI.SetProperty "fold.at.else", "1"
    Else
      DirectSCI.SetProperty "fold.at.else", "0"
    End If
End Property

Public Property Get FoldComment() As Boolean
    FoldComment = m_FoldComment
End Property

Public Property Let FoldComment(ByVal New_FoldComment As Boolean)
    m_FoldComment = New_FoldComment
    PropertyChanged "FoldComment"
    If FoldComment = True Then
      DirectSCI.SetProperty "fold.comment", "1"
    Else
      DirectSCI.SetProperty "fold.comment", "0"
    End If
End Property

'Public Property Get FoldCompact() As Boolean
'    FoldCompact = m_FoldCompact
'End Property
'
'Public Property Let FoldCompact(ByVal New_Compact As Boolean)
'    m_FoldCompact = New_Compact
'    PropertyChanged "FoldComment"
'    If FoldCompact = True Then
'      DirectSCI.SetProperty "fold.compact", "1"
'    Else
'      DirectSCI.SetProperty "fold.compact", "0"
'    End If
'End Property

'Public Property Get FoldHTML() As Boolean
'    FoldHTML = m_FoldHTML
'End Property
'
'Public Property Let FoldHTML(ByVal New_FoldHTML As Boolean)
'    m_FoldHTML = New_FoldHTML
'    PropertyChanged "FoldHTML"
'    If FoldHTML = True Then
'      DirectSCI.SetProperty "fold.HTML", "1"
'    Else
'      DirectSCI.SetProperty "fold.HTML", "0"
'    End If
'End Property

Public Sub FoldAll()
  Dim MaxLine As Long, LineSeek As Long
  MaxLine = DirectSCI.GetLineCount
  DirectSCI.Colourise 0, -1
  For LineSeek = 0 To MaxLine - 1
    If DirectSCI.GetFoldLevel(LineSeek) And SC_FOLDLEVELHEADERFLAG Then
      DirectSCI.ToggleFold LineSeek
    End If
  Next
  DirectSCI.ShowLines 0, 0
End Sub

'===============================[ end folding ] =======================================


'================================[ markers ] ==============================
Private Sub DefineMarker(marknum As Long, Marker As Long)
  Call DirectSCI.MarkerDefine(marknum, Marker)
End Sub

Public Sub ToggleMarker(Optional line As Long = -1)
Attribute ToggleMarker.VB_Description = "Toggels the marker for the specified line. By default uses currently active line."
  On Error Resume Next
  If line = -1 Then line = CurrentLine
  If GetMarker(line) = 4 Then
        DeleteMarker line, 2
  Else
        SetMarker line, 2
  End If
End Sub

Private Function GetMarker(iLine As Long) As Long
    GetMarker = SendEditor(SCI_MARKERGET, iLine)
End Function

Public Sub DeleteMarker(iLine As Long, Optional marknum As Long = 2)
     Dim i As Single
     For i = 0 To 5 'weird bug
        SendEditor SCN_MARKERDELETE, iLine, marknum
     Next
End Sub

Public Sub NextMarker(lline As Long, Optional marknum As Long = 2)
Attribute NextMarker.VB_Description = "Goes to the next marker in document after line argument. MarkNum is the marker group?"
  Dim x As Long
  x = SendEditor(SCN_MARKERNEXT, lline, marknum)
  If x = -1 Then
        x = SendEditor(SCN_MARKERNEXT, 0, marknum)
  End If
  DirectSCI.GotoLine x
End Sub

Public Sub PrevMarker(lline As Long, Optional marknum As Long = 2)
Attribute PrevMarker.VB_Description = "Goes to previous marker from line. MarkNum is the marker group?"
  Dim x As Long
  x = SendEditor(SCN_MARKERPREVIOUS, lline, marknum)
  If x = -1 Then
        x = SendEditor(SCN_MARKERPREVIOUS, DirectSCI.GetLineCount, marknum)
  End If
  DirectSCI.GotoLine x
End Sub

Public Sub DeleteAllMarkers(Optional marknum As Long = 2)
    SendEditor SCN_MARKERDELETEALL, marknum
End Sub

Public Sub SetMarker(iLine As Long, Optional iMarkerNum As Long = 2)
    SendEditor SCI_MARKERADD, iLine, iMarkerNum
End Sub


Public Sub MarkAll(strFind As String)
      Dim x As Long
      Dim g As Boolean
      Dim bFind As Long
      
      x = DirectSCI.GetCurPos
      DirectSCI.SetSel 0, 0
      Call SendEditor(SCI_SETTARGETSTART, 0)
      Call SendEditor(SCI_SETTARGETEND, DirectSCI.GetTextLength)
      bFind = DirectSCI.SearchInTarget(Len(strFind), strFind)
      g = True
      
      Do While bFind > 0
        
            ' Save some time here.  Since were marking all instances if the same
            ' string is found twice in the same line we don't need to know that.
            ' So once we find it in a line and mark it automaticly jump to the next
            ' line
        
            DirectSCI.GotoPos bFind
            SetMarker CurrentLine, 2
            DirectSCI.GotoLine CurrentLine + 1
            Call SendEditor(SCI_SETTARGETSTART, DirectSCI.GetCurPos)
            Call SendEditor(SCI_SETTARGETEND, DirectSCI.GetTextLength)
            bFind = DirectSCI.SearchInTarget(Len(strFind), strFind)
      Loop
      
      DirectSCI.SetSel x, x
End Sub

'================================[ /markers ] ==============================

'================================[ find/replace stuff ] ==============================
Public Function ReplaceText(strSearchFor As String, _
                            strReplaceWith As String, _
                            Optional ReplaceAll As Boolean = False, _
                            Optional CaseSensative As Boolean = False, _
                            Optional WordStart As Boolean = False, _
                            Optional WholeWord As Boolean = False, _
                            Optional RegExp As Boolean = False _
                ) As Boolean
                
    bRepLng = True
    
    If Find(strSearchFor, 0, True, CaseSensative, WordStart, WholeWord) <> -1 Then
      DirectSCI.ReplaceSel strReplaceWith
      If ReplaceAll Then
            bRepAll = True
            Do Until FindNext() = -1
                 DirectSCI.ReplaceSel strReplaceWith
            Loop
            bRepAll = False
      End If
    End If
    
    bRepLng = False
End Function

Public Function ReplaceAll(strSearchFor As String, _
                           strReplaceWith As String, _
                           Optional CaseSensative As Boolean = False, _
                           Optional WordStart As Boolean = False, _
                           Optional WholeWord As Boolean = False, _
                           Optional RegExp As Boolean = False _
                    ) As Long
Attribute ReplaceAll.VB_Description = "Does not affect current line"
                    
      ReplaceAll = 0
      Dim lval As Long
      Dim lenSearch As Long, lenReplace As Long
      Dim Find As Long
      Dim targetstart As Long, targetend As Long, pos As Long, docLen As Long
      
      If strSearchFor = "" Then Exit Function
      
      lval = 0
      If CaseSensative Then lval = lval Or SCFIND_MATCHCASE
      If WordStart Then lval = lval Or SCFIND_WORDSTART
      If WholeWord Then lval = lval Or SCFIND_WHOLEWORD
      If RegExp Then lval = lval Or SCFIND_REGEXP
      
      targetstart = 0
      docLen = DirectSCI.GetTextLength
      lenSearch = Len(strSearchFor)
      lenReplace = Len(strReplaceWith)
    
      targetend = docLen
      Call SendEditor(SCI_SETSEARCHFLAGS, lval)
      Call SendEditor(SCI_SETTARGETSTART, targetstart)
      Call SendEditor(SCI_SETTARGETEND, targetend)
      Find = SendMessageString(SCI, SCI_SEARCHINTARGET, lenSearch, strSearchFor)
      
      Do Until Find = -1
            targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
            targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
            DirectSCI.ReplaceTarget lenReplace, strReplaceWith
            targetstart = targetstart + lenReplace
            targetend = docLen
            ReplaceAll = ReplaceAll + 1
            Call SendEditor(SCI_SETTARGETSTART, targetstart)
            Call SendEditor(SCI_SETTARGETEND, targetend)
            Find = SendMessageString(SCI, SCI_SEARCHINTARGET, lenSearch, strSearchFor)
      Loop
      
End Function

Public Function FindAll(sSearch As String, _
                        ByRef out_StartPositions() As Long, _
                        Optional ByVal searchSelectionOnly As Boolean, _
                        Optional CaseSensative As Boolean = False, _
                        Optional WordStart As Boolean = False, _
                        Optional WholeWord As Boolean = False, _
                        Optional RegExp As Boolean = False _
                ) As Long
Attribute FindAll.VB_Description = "Returns number of indexes added to the out_StartPositions array or -1 on failure"
                
     Dim ret() As Long
     Dim x As Long
     Dim hits As Long
     
     hits = -1
     x = Find(sSearch, 0, False, CaseSensative, WordStart, WholeWord, RegExp)
     
     Do While x <> -1
        hits = hits + 1
        push out_StartPositions, x
        x = FindNext()
     Loop
         
     FindAll = hits
                
End Function

Public Function Find(sSearch As String, _
                        Optional startPos As Long = 0, _
                        Optional autoSelect As Boolean = True, _
                        Optional CaseSensative As Boolean = False, _
                        Optional WordStart As Boolean = False, _
                        Optional WholeWord As Boolean = False, _
                        Optional RegExp As Boolean = False _
                ) As Long
                
    Dim lval As Long, result As Long, targetstart As Long, targetend As Long
    
    ' Sending a null string to scintilla for the find text will cause errors!
    If sSearch = "" Then Exit Function
    
    'these are used in findNext
    bFindEvent = True
    bAutoSelectFinds = autoSelect
    bCase = CaseSensative
    bWholeWord = WholeWord
    bRegEx = RegExp
    bWordStart = WordStart
    strFind = sSearch
    
    lval = 0
    If CaseSensative Then lval = lval Or SCFIND_MATCHCASE
    If WordStart Then lval = lval Or SCFIND_WORDSTART
    If WholeWord Then lval = lval Or SCFIND_WHOLEWORD
    If RegExp Then lval = lval Or SCFIND_REGEXP
    
    Call SendEditor(SCI_SETSEARCHFLAGS, lval)
    Call SendEditor(SCI_SETTARGETSTART, startPos)
    Call SendEditor(SCI_SETTARGETEND, DirectSCI.GetTextLength)
    result = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(sSearch), sSearch)
    
    If autoSelect And result > -1 Then
        targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
        targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))      'for regex endpos != len(txttofind)
        DirectSCI.SetSel targetstart, targetend
    End If
    
    Find = result
    LastFindPos = result + 1
    
End Function

Public Function FindNext(Optional wrap As Boolean = False) As Long
  'If no find events have occurred exit this sub or it may cause errors.
  If bFindEvent = False Then Exit Function
  If wrap And LastFindPos >= DirectSCI.GetTextLength Then LastFindPos = 0
  FindNext = Find(strFind, LastFindPos, bAutoSelectFinds, bCase, bWordStart, bWholeWord, bRegEx)
End Function

Public Function ShowFindReplace() As Object
  On Error Resume Next
  'Load frmReplace
  frmReplace.LaunchReplaceForm Me
  Set frmReplace.Icon = UserControl.Parent.Icon
  Set ShowFindReplace = frmReplace
End Function

'================================[ /find replace stuff ] ==============================

Public Sub ShowAbout()
    On Error Resume Next
    Load frmAbout
    Set frmAbout.Icon = UserControl.Parent.Icon
    frmAbout.LaunchForm Me
    frmAbout.show vbModal
    Unload frmAbout
End Sub

Public Sub ShowGoto()
    On Error Resume Next
    Dim sline As Long
    Dim line As Long
    sline = Trim(InputBox("Goto Line:"))
    If Len(sline) <> 0 Then
        line = CLng(sline)
        If Err.Number = 0 Then Me.GotoLineCentered line
    End If
End Sub

'==============================[ private functions ]===================================

Private Function ToLastSpaceCount() As Long
  ' This function will figure out how many characters there are in the currently
  ' selected word.  It gets the line text, finds the position of the caret in
  ' the line text, then converts the line to a byte array to do a faster compare
  ' till it reaches something not interpreted as a letter IE a space or a
  ' line break.  This is kind of overly complex but seems to be faster overall

  Dim L As Long, i As Long, current As Long, pos As Long, startWord As Long, iHold As Long
  Dim str As String, bByte() As Byte, strTmp As String
  Dim line As String
  line = GetLineText(CurrentLine)
  current = GetCaretInLine
  startWord = current

  Str2Byte line, bByte()

  iHold = 0
  While (startWord > 0) And InStr(1, CallTipWordCharacters, strTmp) > 0
    startWord = startWord - 1
    iHold = iHold + 1
    If startWord >= 0 Then
      strTmp = Chr(bByte(startWord))
    End If
  Wend
  If strTmp = " " Or strTmp = "." Then iHold = iHold - 1
  ToLastSpaceCount = iHold
End Function


Private Sub MaintainIndent()
  On Error Resume Next
  Dim g As Long
  Dim indentAmount As Long
  Dim lastLine As Long
  Dim curLine As Long
  g = DirectSCI.GetCurPos
  ' Get the current line
  curLine = CurrentLine + 1
  ' Get the previous line
  lastLine = curLine - 1

  If GetLineText(lastLine - 1) = "" Then
    'We can move on because in this case there is no text on the
    'previous line.
    Exit Sub
  End If
  indentAmount = 0
  While (lastLine >= 0) And (DirectSCI.GetLineEndPosition(lastLine) - PositionFromLine(lastLine) = 0)
    ' Loop threw the line counting spaces
    lastLine = lastLine - 1
    If lastLine >= 0 Then
      indentAmount = DirectSCI.GetLineIndentation(lastLine)
    End If
    If indentAmount > 0 Then
      Call DirectSCI.SetLineIndentation(curLine - 1, indentAmount)
      Call SetCurrentPosition(GetLineIndentPosition(curLine - 1))

      Call DirectSCI.SetSel(DirectSCI.GetCurPos, DirectSCI.GetCurPos)
    End If
  Wend
End Sub


' This function is utilized to return the modified position of the
' mousecursor on a window
Private Function GetWindowCursorPos(Window As Long) As POINTAPI
  Dim lP As POINTAPI
  Dim rct As RECT
  GetCursorPos lP
  GetWindowRect Window, rct
  GetWindowCursorPos.x = lP.x - rct.Left
  If GetWindowCursorPos.x < 0 Then GetWindowCursorPos.x = 0
  GetWindowCursorPos.y = lP.y - rct.Top
  If GetWindowCursorPos.y < 0 Then GetWindowCursorPos.y = 0
End Function

Private Sub RemoveHotKeys()
  ' This just removes some of the common hot keys that
  ' could cause scintilla to interfere with the application
  
  'apparent the sci hot keys arent reliable? - we will do it ourselves in the hookproc
  DirectSCI.ClearCmdKey Asc("A") + LShift(SCMOD_CTRL, 16) 'sel all
  DirectSCI.ClearCmdKey Asc("V") + LShift(SCMOD_CTRL, 16) 'paste
  DirectSCI.ClearCmdKey Asc("X") + LShift(SCMOD_CTRL, 16) 'cut
  DirectSCI.ClearCmdKey Asc("Z") + LShift(SCMOD_CTRL, 16) 'undo
  
  DirectSCI.ClearCmdKey Asc("Y") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("W") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("B") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("C") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("D") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("E") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("F") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("G") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("H") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("I") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("J") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("K") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("L") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("M") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("N") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("O") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("P") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("Q") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("R") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("S") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("T") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("U") + LShift(SCMOD_CTRL, 16)
   'AssignCmdKey 32 + LShift(SCMOD_CTRL, 16), SCI_AUTOCSHOW
End Sub

Public Function WordUnderMouse(pos As Long, Optional ignoreWhiteSpace As Boolean = False) As String
    Dim sWord As Long, eWord As Long
    
    On Error Resume Next
    'behavior warning.. space characters are counted as words we should count space chars
    'back from pos and pos -= spaceCount
    
    If ignoreWhiteSpace Then pos = pos - GetSpaceCountBack(pos)
    
    
    sWord = DirectSCI.WordStartPosition(pos, True) + 1
    eWord = DirectSCI.WordEndPosition(pos, True) + 1
    WordUnderMouse = Mid(Me.Text, sWord, eWord - sWord)

End Function

'gets the number of spaces counting back to next non white space character
Private Function GetSpaceCountBack(pos As Long)
    On Error Resume Next
    Dim lline As Long, curpos As Long, lText As String, i As Long
    Dim lStart As Long, curCol As Long, b() As Byte, count As Long
    
    lline = DirectSCI.LineFromPosition(pos)
    lText = GetLineText(lline)
    lStart = PositionFromLine(lline)
    curCol = pos - lStart
    
    lText = Left(lText, curCol)
    If Len(lText) = 0 Then Exit Function
    
    b() = StrConv(lText, vbFromUnicode)
    For i = UBound(b) To 0 Step -1
        If b(i) = &H20 Or b(i) = 9 Then
            count = count + 1
        Else
            Exit For
        End If
    Next
    
    GetSpaceCountBack = count
    
End Function

Sub LockEditor(Optional locked As Boolean = True)
    Dim i As Long
    
    If locked Then
        ReadOnly = True
        DirectSCI.StyleSetBack 32, &HF0F0F0
        For i = 0 To 127
            DirectSCI.StyleSetBack i, &HF0F0F0
        Next i
    Else
        ReadOnly = False
        SetHighlighter currentHighlighter
    End If
    
End Sub

Function isMouseOverCallTip() As Boolean
    Dim p As POINTAPI
    Dim hWin As Long
    Dim sz As Long
    Dim sClassName As String * 100
    
    On Error Resume Next
    
    GetCursorPos p
    hWin = WindowFromPoint(p.x, p.y)
    sz = GetClassName(hWin, sClassName, 100)
    If Left(sClassName, sz) = "CallTip" Then isMouseOverCallTip = True
    
End Function





'these would be for setting properties through the IDE to default values per instance, just do it through code
'most defaults are good for me..

''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    m_AutoCloseBraces = PropBag.ReadProperty("AutoCloseBraces", m_def_AutoCloseBraces)
'    m_AutoCloseQuotes = PropBag.ReadProperty("AutoCloseQuotes", m_def_AutoCloseQuotes)
'    m_TabIndents = PropBag.ReadProperty("TabIndents", m_def_TabIndents)
'    m_BackSpaceUnIndents = PropBag.ReadProperty("BackSpaceUnIndents", m_def_BackSpaceUnIndents)
'    m_BookmarkBack = PropBag.ReadProperty("BookmarkBack", m_def_BookmarkBack)
'    m_BookMarkFore = PropBag.ReadProperty("BookMarkFore", m_def_BookMarkFore)
'    m_MarkerBack = PropBag.ReadProperty("MarkerBack", m_def_MarkerBack)
'    m_MarkerFore = PropBag.ReadProperty("MarkerFore", m_def_MarkerFore)
'    m_TabWidth = PropBag.ReadProperty("TabWidth", m_def_TabWidth)
'    m_CaretForeColor = PropBag.ReadProperty("CaretForeColor", m_def_CaretForeColor)
'    m_CaretWidth = PropBag.ReadProperty("CaretWidth", m_def_CaretWidth)
'    m_EdgeColor = PropBag.ReadProperty("EdgeColor", m_def_EdgeColor)
'    m_EOLMode = PropBag.ReadProperty("EOLMode", m_def_EOLMode)
'    m_HighlightBraces = PropBag.ReadProperty("HighlightBraces", m_def_HighlightBraces)
'    m_ClearUndoAfterSave = PropBag.ReadProperty("ClearUndoAfterSave", m_def_ClearUndoAfterSave)
'    m_EndAtLastLine = PropBag.ReadProperty("EndAtLastLine", m_def_EndAtLastLine)
'    m_MaintainIndentation = PropBag.ReadProperty("MaintainIndentation", m_def_MaintainIndentation)
'    m_OverType = PropBag.ReadProperty("OverType", m_def_OverType)
'
'    m_AutoCloseBraces = PropBag.ReadProperty("AutoCloseBraces", m_def_AutoCloseBraces)
'    m_AutoCloseQuotes = PropBag.ReadProperty("AutoCloseQuotes", m_def_AutoCloseQuotes)
'    m_BraceHighlight = PropBag.ReadProperty("BraceHighlight", m_def_BraceHighlight)
'    m_CaretForeColor = PropBag.ReadProperty("CaretForeColor", m_def_CaretForeColor)
'    m_LineBackColor = PropBag.ReadProperty("LineBackColor", m_def_LineBackColor)
'    m_LineVisible = PropBag.ReadProperty("LineVisible", m_def_LineVisible)
'    m_CaretWidth = PropBag.ReadProperty("CaretWidth", m_def_CaretWidth)
'    m_ClearUndoAfterSave = PropBag.ReadProperty("ClearUndoAfterSave", m_def_ClearUndoAfterSave)
'    m_BookmarkBack = PropBag.ReadProperty("BookMarkBack", m_def_BookmarkBack)
'    m_BookMarkFore = PropBag.ReadProperty("BookMarkFore", m_def_BookMarkFore)
'    m_FoldHi = PropBag.ReadProperty("FoldHi", m_def_FoldHi)
'    m_FoldLo = PropBag.ReadProperty("FoldLo", m_def_FoldLo)
'    m_MarkerBack = PropBag.ReadProperty("MarkerBack", m_def_MarkerBack)
'    m_MarkerFore = PropBag.ReadProperty("MarkerFore", m_def_MarkerFore)
'    m_SelBack = PropBag.ReadProperty("SelBack", m_def_SelBack)
'    m_SelFore = PropBag.ReadProperty("SelFore", m_def_SelFore)
'    m_EndAtLastLine = PropBag.ReadProperty("EndAtLastLine", m_def_EndAtLastLine)
'    m_OverType = PropBag.ReadProperty("OverType", m_def_OverType)
'    m_ScrollBarH = PropBag.ReadProperty("ScrollBarH", m_def_ScrollBarH)
'    m_ScrollBarV = PropBag.ReadProperty("ScrollBarV", m_def_ScrollBarV)
'    m_ViewEOL = PropBag.ReadProperty("ViewEOL", m_def_ViewEOL)
'    m_ViewWhiteSpace = PropBag.ReadProperty("ViewWhiteSpace", m_def_ViewWhiteSpace)
'    m_ShowCallTips = PropBag.ReadProperty("ShowCallTips", m_def_ShowCallTips)
'    bShowCallTips = m_ShowCallTips
'    m_EdgeColor = PropBag.ReadProperty("EdgeColor", m_def_EdgeColor)
'    m_EdgeColumn = PropBag.ReadProperty("EdgeColumn", m_def_EdgeColumn)
'    m_EdgeMode = PropBag.ReadProperty("EdgeMode", m_def_EdgeMode)
'    m_EOL = PropBag.ReadProperty("EOL", m_def_EOL)
'    m_Folding = PropBag.ReadProperty("Folding", m_def_Folding)
'    m_Gutter0Type = PropBag.ReadProperty("Gutter0Type", m_def_Gutter0Type)
'    m_Gutter0Width = PropBag.ReadProperty("Gutter0Width", m_def_Gutter0Width)
'    m_Gutter1Type = PropBag.ReadProperty("Gutter1Type", m_def_Gutter1Type)
'    m_Gutter1Width = PropBag.ReadProperty("Gutter1Width", m_def_Gutter1Width)
'    m_Gutter2Type = PropBag.ReadProperty("Gutter2Type", m_def_Gutter2Type)
'    m_Gutter2Width = PropBag.ReadProperty("Gutter2Width", m_def_Gutter2Width)
'    m_MaintainIndentation = PropBag.ReadProperty("MaintainIndentation", m_def_MaintainIndentation)
'    m_TabIndents = PropBag.ReadProperty("TabIndents", m_def_TabIndents)
'    m_BackSpaceUnIndents = PropBag.ReadProperty("BackSpaceUnIndents", m_def_BackSpaceUnIndents)
'    m_IndentWidth = PropBag.ReadProperty("IndentWidth", m_def_IndentWidth)
'    m_UseTabs = PropBag.ReadProperty("UseTabs", m_def_UseTabs)
'    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
'    m_FoldMarker = PropBag.ReadProperty("FoldMarker", m_def_FoldMarker)
'    m_MarginFore = PropBag.ReadProperty("MarginFore", m_def_MarginFore)
'    m_MarginBack = PropBag.ReadProperty("MarginBack", m_def_MarginBack)
'    m_Text = PropBag.ReadProperty("Text", m_def_Text)
'    m_SelText = PropBag.ReadProperty("SelText", m_def_SelText)
'    m_AutoCompleteStart = PropBag.ReadProperty("AutoCompleteStart", m_def_AutoCompleteStart)
'    m_AutoCompleteOnCTRLSpace = PropBag.ReadProperty("AutoCompleteOnCTRLSpace", m_def_AutoCompleteOnCTRLSpace)
'    m_AutoCompleteString = PropBag.ReadProperty("AutoCompleteString", m_def_AutoCompleteString)
'    m_AutoShowAutoComplete = PropBag.ReadProperty("AutoShowAutoComplete", m_def_AutoShowAutoComplete)
'    m_ContextMenu = PropBag.ReadProperty("ContextMenu", m_def_ContextMenu)
'    m_IgnoreAutoCompleteCase = PropBag.ReadProperty("IgnoreAutoCompleteCase", m_def_IgnoreAutoCompleteCase)
'    m_LineNumbers = PropBag.ReadProperty("LineNumbers", m_def_LineNumbers)
'    m_ReadOnly = PropBag.ReadProperty("ReadOnly", m_def_ReadOnly)
'    m_ScrollWidth = PropBag.ReadProperty("ScrollWidth", m_def_ScrollWidth)
'    m_ShowFlags = PropBag.ReadProperty("ShowFlags", m_def_ShowFlags)
'    m_FoldAtElse = PropBag.ReadProperty("FoldAtElse", m_def_FoldAtElse)
'    m_FoldComment = PropBag.ReadProperty("FoldComment", m_def_FoldComment)
'    m_FoldCompact = PropBag.ReadProperty("FoldCompact", m_def_FoldCompact)
'    m_FoldHTML = PropBag.ReadProperty("FoldHTML", m_def_FoldHTML)
'    m_IndentationGuide = PropBag.ReadProperty("IndentationGuide", m_def_IndentationGuide)
'    m_SelStart = PropBag.ReadProperty("SelStart", m_def_SelStart)
'    m_SelEnd = PropBag.ReadProperty("SelEnd", m_def_SelEnd)
'    m_BraceMatch = PropBag.ReadProperty("BraceMatch", m_def_BraceMatch)
'    m_BraceBad = PropBag.ReadProperty("BraceBad", m_def_BraceBad)
'    m_BraceMatchBold = PropBag.ReadProperty("BraceMatchBold", m_def_BraceMatchBold)
'    m_BraceMatchItalic = PropBag.ReadProperty("BraceMatchItalic", m_def_BraceMatchItalic)
'    m_BraceMatchUnderline = PropBag.ReadProperty("BraceMatchUnderline", m_def_BraceMatchUnderline)
'    m_BraceMatchBack = PropBag.ReadProperty("BraceMatchBack", m_def_BraceMatchBack)
'    m_BraceBadBack = PropBag.ReadProperty("BraceBadBack", m_def_BraceBadBack)
'    m_CodePage = PropBag.ReadProperty("CodePage", m_def_CodePage)
'
'
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Exit Sub
'
'    Call PropBag.WriteProperty("AutoCloseBraces", m_AutoCloseBraces, m_def_AutoCloseBraces)
'    Call PropBag.WriteProperty("AutoCloseQuotes", m_AutoCloseQuotes, m_def_AutoCloseQuotes)
'    Call PropBag.WriteProperty("TabIndents", m_TabIndents, m_def_TabIndents)
'    Call PropBag.WriteProperty("BackSpaceUnIndents", m_BackSpaceUnIndents, m_def_BackSpaceUnIndents)
'    Call PropBag.WriteProperty("BookmarkBack", m_BookmarkBack, m_def_BookmarkBack)
'    Call PropBag.WriteProperty("BookMarkFore", m_BookMarkFore, m_def_BookMarkFore)
'    Call PropBag.WriteProperty("MarkerBack", m_MarkerBack, m_def_MarkerBack)
'    Call PropBag.WriteProperty("MarkerFore", m_MarkerFore, m_def_MarkerFore)
'    Call PropBag.WriteProperty("TabWidth", m_TabWidth, m_def_TabWidth)
'    Call PropBag.WriteProperty("CaretForeColor", m_CaretForeColor, m_def_CaretForeColor)
'    Call PropBag.WriteProperty("CaretWidth", m_CaretWidth, m_def_CaretWidth)
'    Call PropBag.WriteProperty("EdgeColor", m_EdgeColor, m_def_EdgeColor)
'    Call PropBag.WriteProperty("EOLMode", m_EOLMode, m_def_EOLMode)
'    Call PropBag.WriteProperty("HighlightBraces", m_HighlightBraces, m_def_HighlightBraces)
'    Call PropBag.WriteProperty("ClearUndoAfterSave", m_ClearUndoAfterSave, m_def_ClearUndoAfterSave)
'    Call PropBag.WriteProperty("EndAtLastLine", m_EndAtLastLine, m_def_EndAtLastLine)
'    Call PropBag.WriteProperty("MaintainIndentation", m_MaintainIndentation, m_def_MaintainIndentation)
'    Call PropBag.WriteProperty("OverType", m_OverType, m_def_OverType)
'
'    Call PropBag.WriteProperty("AutoCloseBraces", m_AutoCloseBraces, m_def_AutoCloseBraces)
'    Call PropBag.WriteProperty("AutoCloseQuotes", m_AutoCloseQuotes, m_def_AutoCloseQuotes)
'    Call PropBag.WriteProperty("BraceHighlight", m_BraceHighlight, m_def_BraceHighlight)
'    Call PropBag.WriteProperty("CaretForeColor", m_CaretForeColor, m_def_CaretForeColor)
'    Call PropBag.WriteProperty("LineBackColor", m_LineBackColor, m_def_LineBackColor)
'    Call PropBag.WriteProperty("LineVisible", m_LineVisible, m_def_LineVisible)
'    Call PropBag.WriteProperty("CaretWidth", m_CaretWidth, m_def_CaretWidth)
'    Call PropBag.WriteProperty("ClearUndoAfterSave", m_ClearUndoAfterSave, m_def_ClearUndoAfterSave)
'    Call PropBag.WriteProperty("BookMarkBack", m_BookmarkBack, m_def_BookmarkBack)
'    Call PropBag.WriteProperty("BookMarkFore", m_BookMarkFore, m_def_BookMarkFore)
'    Call PropBag.WriteProperty("FoldHi", m_FoldHi, m_def_FoldHi)
'    Call PropBag.WriteProperty("FoldLo", m_FoldLo, m_def_FoldLo)
'    Call PropBag.WriteProperty("MarkerBack", m_MarkerBack, m_def_MarkerBack)
'    Call PropBag.WriteProperty("MarkerFore", m_MarkerFore, m_def_MarkerFore)
'    Call PropBag.WriteProperty("SelBack", m_SelBack, m_def_SelBack)
'    Call PropBag.WriteProperty("SelFore", m_SelFore, m_def_SelFore)
'    Call PropBag.WriteProperty("EndAtLastLine", m_EndAtLastLine, m_def_EndAtLastLine)
'    Call PropBag.WriteProperty("OverType", m_OverType, m_def_OverType)
'    Call PropBag.WriteProperty("ScrollBarH", m_ScrollBarH, m_def_ScrollBarH)
'    Call PropBag.WriteProperty("ScrollBarV", m_ScrollBarV, m_def_ScrollBarV)
'    Call PropBag.WriteProperty("ViewEOL", m_ViewEOL, m_def_ViewEOL)
'    Call PropBag.WriteProperty("ViewWhiteSpace", m_ViewWhiteSpace, m_def_ViewWhiteSpace)
'    Call PropBag.WriteProperty("ShowCallTips", m_ShowCallTips, m_def_ShowCallTips)
'    Call PropBag.WriteProperty("EdgeColor", m_EdgeColor, m_def_EdgeColor)
'    Call PropBag.WriteProperty("EdgeColumn", m_EdgeColumn, m_def_EdgeColumn)
'    Call PropBag.WriteProperty("EdgeMode", m_EdgeMode, m_def_EdgeMode)
'    Call PropBag.WriteProperty("EOL", m_EOL, m_def_EOL)
'    Call PropBag.WriteProperty("Folding", m_Folding, m_def_Folding)
'    Call PropBag.WriteProperty("Gutter0Type", m_Gutter0Type, m_def_Gutter0Type)
'    Call PropBag.WriteProperty("Gutter0Width", m_Gutter0Width, m_def_Gutter0Width)
'    Call PropBag.WriteProperty("Gutter1Type", m_Gutter1Type, m_def_Gutter1Type)
'    Call PropBag.WriteProperty("Gutter1Width", m_Gutter1Width, m_def_Gutter1Width)
'    Call PropBag.WriteProperty("Gutter2Type", m_Gutter2Type, m_def_Gutter2Type)
'    Call PropBag.WriteProperty("Gutter2Width", m_Gutter2Width, m_def_Gutter2Width)
'    Call PropBag.WriteProperty("MaintainIndentation", m_MaintainIndentation, m_def_MaintainIndentation)
'    Call PropBag.WriteProperty("TabIndents", m_TabIndents, m_def_TabIndents)
'    Call PropBag.WriteProperty("BackSpaceUnIndents", m_BackSpaceUnIndents, m_def_BackSpaceUnIndents)
'    Call PropBag.WriteProperty("IndentWidth", m_IndentWidth, m_def_IndentWidth)
'    Call PropBag.WriteProperty("UseTabs", m_UseTabs, m_def_UseTabs)
'    Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
'    Call PropBag.WriteProperty("FoldMarker", m_FoldMarker, m_def_FoldMarker)
'    Call PropBag.WriteProperty("MarginFore", m_MarginFore, m_def_MarginFore)
'    Call PropBag.WriteProperty("MarginBack", m_MarginBack, m_def_MarginBack)
'    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
'    Call PropBag.WriteProperty("SelText", m_SelText, m_def_SelText)
'    Call PropBag.WriteProperty("AutoCompleteStart", m_AutoCompleteStart, m_def_AutoCompleteStart)
'    Call PropBag.WriteProperty("AutoCompleteOnCTRLSpace", m_AutoCompleteOnCTRLSpace, m_def_AutoCompleteOnCTRLSpace)
'    Call PropBag.WriteProperty("AutoCompleteString", m_AutoCompleteString, m_def_AutoCompleteString)
'    Call PropBag.WriteProperty("AutoShowAutoComplete", m_AutoShowAutoComplete, m_def_AutoShowAutoComplete)
'    Call PropBag.WriteProperty("ContextMenu", m_ContextMenu, m_def_ContextMenu)
'    Call PropBag.WriteProperty("IgnoreAutoCompleteCase", m_IgnoreAutoCompleteCase, m_def_IgnoreAutoCompleteCase)
'    Call PropBag.WriteProperty("LineNumbers", m_LineNumbers, m_def_LineNumbers)
'    Call PropBag.WriteProperty("ReadOnly", m_ReadOnly, m_def_ReadOnly)
'    Call PropBag.WriteProperty("ScrollWidth", m_ScrollWidth, m_def_ScrollWidth)
'    Call PropBag.WriteProperty("ShowFlags", m_ShowFlags, m_def_ShowFlags)
'    Call PropBag.WriteProperty("FoldAtElse", m_FoldAtElse, m_def_FoldAtElse)
'
'    Call PropBag.WriteProperty("FoldComment", m_FoldComment, m_def_FoldComment)
'    Call PropBag.WriteProperty("FoldCompact", m_FoldCompact, m_def_FoldCompact)
'    Call PropBag.WriteProperty("FoldHTML", m_FoldHTML, m_def_FoldHTML)
'
'    Call PropBag.WriteProperty("IndentationGuide", m_IndentationGuide, m_def_IndentationGuide)
'    Call PropBag.WriteProperty("SelStart", m_SelStart, m_def_SelStart)
'    Call PropBag.WriteProperty("SelEnd", m_SelEnd, m_def_SelEnd)
'    Call PropBag.WriteProperty("BraceMatch", m_BraceMatch, m_def_BraceMatch)
'    Call PropBag.WriteProperty("BraceBad", m_BraceBad, m_def_BraceBad)
'    Call PropBag.WriteProperty("BraceMatchBold", m_BraceMatchBold, m_def_BraceMatchBold)
'    Call PropBag.WriteProperty("BraceMatchItalic", m_BraceMatchItalic, m_def_BraceMatchItalic)
'    Call PropBag.WriteProperty("BraceMatchUnderline", m_BraceMatchUnderline, m_def_BraceMatchUnderline)
'    Call PropBag.WriteProperty("BraceMatchBack", m_BraceMatchBack, m_def_BraceMatchBack)
'    Call PropBag.WriteProperty("BraceBadBack", m_BraceBadBack, m_def_BraceBadBack)
'    Call PropBag.WriteProperty("CodePage", m_CodePage, m_def_CodePage)
'
'End Sub


'
'Public Property Get EOLMode() As EOLStyle
'    EOLMode = m_EOLMode
'End Property
'
'Public Property Let EOLMode(ByVal New_EOLMode As EOLStyle)
'    m_EOLMode = New_EOLMode
'    PropertyChanged "EOLMode"
'End Property
'
'Public Property Get TabWidth() As Long
'    TabWidth = m_TabWidth
'End Property
'
'Public Property Let TabWidth(ByVal New_TabWidth As Long)
'    m_TabWidth = New_TabWidth
'    PropertyChanged "TabWidth"
'End Property
'
'Public Property Get HighlightBraces() As Boolean    'When set to true any braces the cursor is next to will be highlighted.
'    HighlightBraces = m_BraceHighlight
'End Property
'
'Public Property Let HighlightBraces(ByVal New_BraceHighlight As Boolean)
'    m_BraceHighlight = New_BraceHighlight
'    PropertyChanged "BraceHighlight"
'End Property
'
'Public Property Get CaretForeColor() As OLE_COLOR   'Set's the color of the caret.
'    CaretForeColor = m_CaretForeColor
'End Property
'
'Public Property Get CaretWidth() As Long    'Allow's you to control the width of the caret line.  The maximum value is 3.
'    CaretWidth = m_CaretWidth
'End Property
'
'Public Property Let CaretWidth(ByVal New_CaretWidth As Long)
'    If New_CaretWidth > 3 Then New_CaretWidth = 3
'    m_CaretWidth = New_CaretWidth
'    PropertyChanged "CaretWidth"
'    DirectSCI.SetCaretWidth m_CaretWidth
'End Property
'
'Public Property Get ClearUndoAfterSave() As Boolean 'If set to true then the undo buffer will be cleared when calling SaveToFile.
'    ClearUndoAfterSave = m_ClearUndoAfterSave
'End Property
'
'Public Property Let ClearUndoAfterSave(ByVal New_ClearUndoAfterSave As Boolean)
'    m_ClearUndoAfterSave = New_ClearUndoAfterSave
'    PropertyChanged "ClearUndoAfterSave"
'End Property
'
'Public Property Get EndAtLastLine() As Boolean  'If set to true then the document won't scroll past the last line.  If false it will allow you to scroll a bit past the end of the file.
'    EndAtLastLine = m_EndAtLastLine
'End Property
'
'Public Property Let EndAtLastLine(ByVal New_EndAtLastLine As Boolean)
'    m_EndAtLastLine = New_EndAtLastLine
'    PropertyChanged "EndAtLastLine"
'    DirectSCI.SetEndAtLastLine m_EndAtLastLine
'End Property
'
'Public Property Get OverType() As Boolean   'If true then entered text will overtype any text beyond it.
'    OverType = m_OverType
'End Property
'
'Public Property Let OverType(ByVal New_OverType As Boolean)
'    m_OverType = New_OverType
'    PropertyChanged "OverType"
'    DirectSCI.SetOvertype m_OverType
'End Property
'
'Public Property Get ScrollBarH() As Boolean  'If true then the horizontal scrollbar will be visible.  If false it will be hidden.
'    ScrollBarH = m_ScrollBarH
'End Property
'
'Public Property Let ScrollBarH(ByVal New_ScrollBarH As Boolean)
'    m_ScrollBarH = New_ScrollBarH
'    PropertyChanged "ScrollBarH"
'    DirectSCI.SetHScrollBar m_ScrollBarH
'End Property
'
'Public Property Get ScrollBarV() As Boolean 'If true then the vertical scrollbar will be visible.  If alse it will be hidden.
'    ScrollBarV = m_ScrollBarV
'End Property
'
'Public Property Let ScrollBarV(ByVal New_ScrollBarV As Boolean)
'    m_ScrollBarV = New_ScrollBarV
'    PropertyChanged "ScrollBarV"
'    DirectSCI.SetVScrollBar New_ScrollBarV
'End Property
'
'Public Property Get ViewEOL() As Boolean    'If this is set to true EOL markers will be displayed.
'    ViewEOL = m_ViewEOL
'End Property
'
'Public Property Let ViewEOL(ByVal New_ViewEOL As Boolean)
'    m_ViewEOL = New_ViewEOL
'    PropertyChanged "ViewEOL"
'    DirectSCI.SetViewEOL New_ViewEOL
'End Property
'
'Public Property Get EdgeColor() As OLE_COLOR 'This allows you to control the color of the Edge line.
'    EdgeColor = m_EdgeColor
'End Property
'
'Public Property Let EdgeColor(ByVal New_EdgeColor As OLE_COLOR)
'    m_EdgeColor = New_EdgeColor
'    PropertyChanged "EdgeColor"
'    DirectSCI.SetEdgeColour m_EdgeColor
'End Property
'
'Public Property Get EdgeColumn() As Long    'This allows you to control which column the edge line is located at.
'    EdgeColumn = m_EdgeColumn
'End Property
'
'Public Property Let EdgeColumn(ByVal New_EdgeColumn As Long)
'    m_EdgeColumn = New_EdgeColumn
'    PropertyChanged "EdgeColumn"
'    DirectSCI.SetEdgeColumn m_EdgeColumn
'End Property
'
'Public Property Get EdgeMode() As edge  'This allow's you to control which edge mode to utilize.
'    EdgeMode = m_EdgeMode
'End Property
'
'Public Property Let EdgeMode(ByVal New_EdgeMode As edge)
'    m_EdgeMode = New_EdgeMode
'    PropertyChanged "EdgeMode"
'    DirectSCI.SetEdgeMode m_EdgeMode
'End Property
'
'Public Property Get EOL() As EOLStyle   'This allows you to control which EOL style to utilize.  Scintilla supports CR+LF, CR, and LF.
'    EOL = m_EOL
'End Property
'
'Public Property Let EOL(ByVal New_EOL As EOLStyle)
'    m_EOL = New_EOL
'    PropertyChanged "EOL"
'    DirectSCI.SetEOLMode m_EOL
'End Property
'
'Public Property Get ViewWhiteSpace() As Boolean 'When this is set to true whitespace markers will be visible.
'    ViewWhiteSpace = m_ViewWhiteSpace
'End Property
'
'Public Property Let ViewWhiteSpace(ByVal New_ViewWhiteSpace As Boolean)
'    m_ViewWhiteSpace = New_ViewWhiteSpace
'    PropertyChanged "ViewWhiteSpace"
'    DirectSCI.SetViewWS CLng(m_ViewWhiteSpace)
'End Property
'
'Public Property Get ScrollWidth() As Long   'Scintilla's design does not automatically size the horizontal scrollbar to the size of the longest line.  It gives it a set size.  By default it allows 2000 characters per line.  This allows you to control how far the Horizontal scrollbar can be scrolled.
'    ScrollWidth = m_ScrollWidth
'End Property
'
'Public Property Let ScrollWidth(ByVal New_ScrollWidth As Long)
'    m_ScrollWidth = New_ScrollWidth
'    PropertyChanged "ScrollWidth"
'End Property
'
'Public Function SetSavePoint() As Long
'  DirectSCI.SetSavePoint
'End Function
'
'Public Sub TabRight()
'  SendEditor SCI_TAB
'End Sub
'
'Public Sub TabLeft()
'  SendEditor SCI_BACKTAB
'End Sub

'Public Property Get MarginFore() As OLE_COLOR
'    MarginFore = m_MarginFore
'End Property
'
'Public Property Let MarginFore(ByVal New_MarginFore As OLE_COLOR)
'    m_MarginFore = New_MarginFore
'    PropertyChanged "MarginFore"
'End Property
'
'Public Property Get MarginBack() As OLE_COLOR
'    MarginBack = m_MarginBack
'End Property
'
'Public Property Let MarginBack(ByVal New_MarginBack As OLE_COLOR)
'    m_MarginBack = New_MarginBack
'    PropertyChanged "MarginBack"
'End Property

'This property is used for the folding gutter's back color.
'The Hi color is the primary color, the Lo Color is the secondary color.
'Public Property Get FoldingGutterColor(primary As Boolean) As OLE_COLOR
'    If primary Then FoldingGutterColor = m_FoldHi Else FoldingGutterColor = mfold_lo
'End Property
'
'Public Property Let FoldingGutterColor(primary As Boolean, ByVal v As OLE_COLOR)
'
'    If primary Then
'        m_FoldHi = v
'        owner.DirectSCI.SetFoldMarginHiColour True, m_FoldHi
'    Else
'        m_FoldLo = v
'        DirectSCI.SetFoldMarginColour True, m_FoldLo
'    End If
'
'End Property


