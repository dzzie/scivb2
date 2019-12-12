VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMuiltiLineTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Sci As SciSimple
Attribute Sci.VB_VarHelpID = -1

Public curLine As Long
Public totalLines As Long
Private tips()

'https://scintilla-interest.narkive.com/3PCaUtF8/calltip-arrows
'
'The text of the calltip is searched for '\001' and '\002' bytes and these
'are displayed as up and down arrows. When the mouse is clicked on the
'calltip, a SCN_CALLTIPCLICK notification is sent to the container with the
'position field set to 1 if the click is in an up arrow, 2 if in a down
'arrow, 0 if elsewhere. The container can use this information to change the
'calltip displayed. Calltips can be redisplayed with new text by calling
'SCI_CALLTIPSHOW rather than requiring a SCI_CALLTIPCANCEL to remove the
'calltip first. This avoids flicker.

Private Property Get curTip() As String
    curTip = curLine & "/" & totalLines & " " & Chr(1) & Chr(2) & " " & tips(curLine)
End Property

Sub InitAndShow(tipAry)
    
    Dim t As String
    Dim i As Long
    
    tips = tipAry
    curLine = 0
    totalLines = UBound(tips)
    Sci.ShowCallTip curTip
    
End Sub

Private Sub Sci_CallTipClick(Position As Long)
    Dim t As String
    
    If Position = 1 Then 'upArrow
        If curLine = 0 Then Exit Sub
        curLine = curLine - 1
    ElseIf Position = 2 Then 'down arrow
        If curLine = totalLines Then Exit Sub
        curLine = curLine + 1
    End If

    Sci.StopCallTip  'without this (which introduces flicker) it fails to update for some randomly
    Sci.ShowCallTip curTip
    

End Sub
