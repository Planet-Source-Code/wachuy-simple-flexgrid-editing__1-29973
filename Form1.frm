VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "FlexGrid"
   ClientHeight    =   4215
   ClientLeft      =   3135
   ClientTop       =   2070
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7185
   Begin VB.TextBox txtFlexGridCell 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   2205
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5985
      TabIndex        =   3
      Top             =   1455
      Width           =   825
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5985
      TabIndex        =   2
      Top             =   1830
      Width           =   825
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3285
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   5794
      _Version        =   393216
      Cols            =   4
      BackColorFixed  =   16744576
      ForeColorFixed  =   12582912
      BackColorBkg    =   16761024
      AllowUserResizing=   3
      FormatString    =   "      | Name                         | Address                  | Tel Num            "
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentRow As Integer
Dim DoNotChange As Boolean

Private Sub cmdAdd_Click()
Dim Answer As Integer
Me.txtFlexGridCell.Visible = False

Answer = MsgBox("At the bottom?", vbYesNo, "Add where...")
Select Case Answer
    Case vbYes
        Me.MSHFlexGrid1.AddItem ""
        CurrentRow = Me.MSHFlexGrid1.Rows - 1 'add bottom of grid
    Case vbNo
        Me.MSHFlexGrid1.AddItem "", 1
End Select

End Sub

Private Sub cmdRemove_Click()
Dim Answer As Integer
Dim ColCounter As Integer
Me.txtFlexGridCell.Visible = False
  Answer = MsgBox("Do you want to remove?", vbYesNo + vbDefaultButton2, "Confirm remove...")
  Select Case Answer
    Case vbYes
        If Me.MSHFlexGrid1.Rows = Me.MSHFlexGrid1.FixedRows + 1 Then
            For ColCounter = 1 To Me.MSHFlexGrid1.Cols - 1
                Me.MSHFlexGrid1.Col = ColCounter
                Me.MSHFlexGrid1.Text = ""
            Next
        Else
            Me.MSHFlexGrid1.RemoveItem CurrentRow
        End If
    Case vbNo
  
  End Select
End Sub


Private Sub MSHFlexGrid1_Click()
'algo to position textbox inside flexgrid cells
'column 0 is used as a marker
Me.txtFlexGridCell.Visible = True

CurrentRow = Me.MSHFlexGrid1.Row

Me.txtFlexGridCell.Height = Me.MSHFlexGrid1.CellHeight - 10 'minus 10 so that grid lines
Me.txtFlexGridCell.Width = Me.MSHFlexGrid1.CellWidth - 10 '  will not be overwritten
Me.txtFlexGridCell.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
Me.txtFlexGridCell.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
DoNotChange = True
Me.txtFlexGridCell.Text = Me.MSHFlexGrid1.Text
DoNotChange = False
Me.txtFlexGridCell.SetFocus
    
End Sub

Private Sub txtFlexGridCell_Change()
    If DoNotChange Then Exit Sub
    Me.MSHFlexGrid1.Text = Me.txtFlexGridCell.Text
End Sub
