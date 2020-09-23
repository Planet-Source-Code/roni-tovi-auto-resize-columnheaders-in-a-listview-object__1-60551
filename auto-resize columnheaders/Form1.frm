VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add New Record"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resize Headers"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column #1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column #2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column #3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Column #4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Column #5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Column #6"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Would you like to enlarge column headers if there will be available space?", vbQuestion + vbYesNo) = vbYes Then
    ResizeLW Me, Me.ListView1
Else
    ResizeLW Me, Me.ListView1, False
End If
End Sub

Private Sub Command2_Click()
Set Item = ListView1.ListItems.Add(, , InputBox("Enter field #1 content"))
For i = 1 To ListView1.ColumnHeaders.Count - 1
Item.SubItems(i) = InputBox("Enter field #" & i + 1 & " content")
Next
End Sub

Private Sub Form_Load()
Set Item = ListView1.ListItems.Add(, , "something in field #1")
Item.SubItems(1) = "something other in field #2"
Item.SubItems(2) = "field #3"
Item.SubItems(3) = "other field #4"
Item.SubItems(4) = "field #5"
Item.SubItems(5) = "#6"
End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Move 0, 0, Me.Width - 100, Me.Height - 1300
Command1.Move Me.Width - Command1.Width - 200, ListView1.Height + 200
Command2.Move 100, Command1.Top
End Sub
