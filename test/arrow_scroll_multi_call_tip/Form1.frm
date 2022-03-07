VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin sci2.SciSimple sci 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   8493
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this sample shows how to use the arrow keys to scroll through multiple call tips
'without having to click on the up and down arrows with the mouse.
'type yy. in the IDE window to trigger the multiline call tip

Implements Project1.iSubclass
Dim multi As New CMuiltiLineTip

Private SC As cSubclass

Private Sub Sci_AutoCompleteEvent(className As String)
    
    Dim prevWord As String
    prevWord = sci.PreviousWord()

    Debug.Print "AutoCompleteEvent: ClassName: " & className & " Prevword: " & prevWord
        
    If className = "yy" Then
         multi.InitAndShow Array("test(xx,yy)", "test(yy)", "test(zz)", "test(kksssssssssssss)")
    End If
        
End Sub

Public Function IsTipActive() As Boolean 'thanks Flauzer!
    Const SCI_CALLTIPACTIVE = 2202
    IsTipActive = sci.DirectSCI.SendEditor(SCI_CALLTIPACTIVE)
End Function

Private Sub Form_Load()
    
    Set multi.sci = sci
    
    Set SC = New cSubclass
    SC.Subclass sci.sciHWND, Me
    SC.AddMsg sci.sciHWND, WM_KEYDOWN, MSG_BEFORE
 
End Sub

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'Debug.Print wParam
    
    If uMsg = WM_KEYDOWN Then
        If IsTipActive() Then
            If wParam = 38 Then 'up arrow
                bHandled = True
                multi.PrevCallTip
            ElseIf wParam = 40 Then 'down arrow
                bHandled = True
                multi.NextCallTip
            End If
        End If
    End If

End Sub
 


 


