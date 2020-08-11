VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Begin VB.Form Form1 
   Caption         =   "Override Goto and Find Dialogs"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   90
      TabIndex        =   1
      Top             =   2430
      Width           =   6405
   End
   Begin sci2.SciSimple sci 
      Height          =   2175
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   3836
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iSubclass

Dim sc As New cSubclass

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_SHIFT = &H10&
Private Const VK_CONTROL = &H11&
Private Const VK_MENU = &H12&  ' Alt key

Private Sub Form_Load()
    
    If sc.Subclass(sci.sciHWND, Me) Then
        sc.AddMsg sci.sciHWND, WM_KEYUP, MSG_BEFORE
        sci.Text = "Hooked, press ctrl-G or ctrl-F to test..."
    Else
        sci.Text = "Failed to subclass sci control"
    End If
    
End Sub

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

    On Error Resume Next
                    
    Select Case uMsg
    
        Case WM_KEYUP
                           
             If piGetShiftState = 4 Then 'CTRL Key
                
                If wParam = Asc("F") Or wParam = Asc("H") Then
                    List1.AddItem "Find replace intercepted"
                    GoTo eatMessage
                End If
                
                If Asc("G") = wParam Then
                    List1.AddItem "Goto intercepted"
                    GoTo eatMessage
                End If
                
            End If
            
    End Select
                    
Exit Sub

eatMessage:
     bHandled = True
     lReturn = 0
     wParam = 0

End Sub

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

