VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form deletedrop 
   BackColor       =   &H00D8C7BC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete/Drop Table"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "Deletedrop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8C7BC&
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      Begin VB.Frame Frame2 
         BackColor       =   &H00D8C7BC&
         Caption         =   "Table Name"
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4215
         Begin VB.CommandButton cmdgo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "GO"
            Default         =   -1  'True
            Height          =   375
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   495
         End
         Begin MSForms.ComboBox cbotables 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   3135
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   8388608
            DisplayStyle    =   3
            Size            =   "5530;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.OptionButton optdrop 
         BackColor       =   &H00D8C7BC&
         Caption         =   "Drop The Table"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optdelete 
         BackColor       =   &H00D8C7BC&
         Caption         =   "Delete All The Records From Table"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   240
         Picture         =   "Deletedrop.frx":0442
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   240
         Picture         =   "Deletedrop.frx":058C
         Top             =   1680
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Delete/Drop Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1965
   End
End
Attribute VB_Name = "deletedrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tablefound       As Boolean

Private Sub cmdgo_Click()
On Error GoTo Jump

If cbotables.Text = "" Then
  MsgBox "Select Table Name", vbExclamation
  Exit Sub
End If

tablefound = False

For i = 0 To cbotables.ListCount - 1
    If Trim(cbotables.Text) = cbotables.List(i) Then
       tablefound = True
       Exit For
    End If
Next

If tablefound = False Then
  MsgBox "No A Valid Table", vbExclamation
  cbotables.Text = ""
  Exit Sub
End If

If optdelete.Value = True Then

   If MsgBox("Delete All The Records from " & cbotables.Text & " ?", vbQuestion + vbYesNo) = vbYes Then
   
       cn.Execute "delete from " & cbotables.Text
       MsgBox "All The Records Are Deleted From " & cbotables.Text, vbInformation
       Temp = cbotables.Text
       FillCombo
       cbotables.Text = Temp
       FillCombo
   
   End If
 
ElseIf optdrop.Value = True Then

   If MsgBox("Drop Table " & cbotables.Text & " ?", vbQuestion + vbYesNo) = vbYes Then
   
      cn.Execute "drop table " & cbotables.Text
      MsgBox cbotables.Text & " Table Is Droped ", vbInformation
      frmmain.Temp = frmmain.cbotables.Text
      frmmain.FillCombo
      frmmain.cbotables.Text = frmmain.Temp
      frmmain.FillGrid
      FillCombo
   
   End If

End If

Exit Sub
Jump:
    
    MsgBox Err.Description, vbCritical
  
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
  FillCombo
End Sub

Public Sub FillCombo()
On Error GoTo Jump

 cbotables.Clear
 cbotables.Text = ""
 
    For Each Table In mCat.Tables
    
     If Table.Type = "TABLE" Then
     
       cbotables.AddItem Table.Name
     
     End If
    
    Next
 
Exit Sub
Jump:
   MsgBox Err.Description, vbCritical

End Sub
