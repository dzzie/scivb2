VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TabStrip ts 
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim c As New Collection
Dim WithEvents sci As SciSimple
Attribute sci.VB_VarHelpID = -1

Private Sub Command1_Click()
    Dim o As Object, oName As String
    Dim t As Tab
    
    oName = "sci" & c.Count
   
    sci.Visible = False
    
    Set o = Controls.Add("sci2.SciSimple", oName, Me)
    Set sci = o
    c.Add o, oName
    sci.Visible = True
    sci.Move 0, ts.Top + ts.Height, Me.ScaleWidth, Me.ScaleHeight
    
    Set t = ts.Tabs.Add(, , oName)
    t.Selected = True
    
End Sub

Private Sub Form_Load()
    Dim o As Object, oName As String
    oName = "sci" & c.Count
    Set o = Controls.Add("sci2.SciSimple", oName, Me)
    Set sci = o
    c.Add o, oName
    ts.Tabs(1).Caption = oName
    sci.Visible = True
    sci.Move 0, ts.Top + ts.Height, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub sci_KeyDown(KeyCode As Long, Shift As Long)
    Debug.Print sci.Name & " keydown: " & KeyCode
End Sub

Private Sub ts_Click()
    Debug.Print "ts.SelectedItem.Index: " & ts.SelectedItem.Index
    Dim o As Object
    sci.Visible = False
    Set o = c(ts.SelectedItem.Index)
    Set sci = o
    sci.Visible = True
End Sub


