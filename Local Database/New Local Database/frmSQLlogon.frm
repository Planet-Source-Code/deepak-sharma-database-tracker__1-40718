VERSION 5.00
Begin VB.Form frmSQLSERVER 
   BackColor       =   &H00D8C7BC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Server"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   Icon            =   "frmSQLlogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSQLlogon.frx":0442
   ScaleHeight     =   2475
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8C7BC&
      Caption         =   "Database "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   3855
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cbosqldatabase 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   240
         Picture         =   "frmSQLlogon.frx":0884
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame frmsql 
      BackColor       =   &H00D8C7BC&
      Caption         =   "Authentication"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtserver 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Text            =   "(local)"
         Top             =   390
         Width           =   2055
      End
      Begin VB.TextBox txtuser 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   870
         Width           =   2055
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1350
         Width           =   2055
      End
      Begin VB.CommandButton cmdconnect 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmSQLlogon.frx":09CE
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User name"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SQL Server"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmSQLSERVER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdconnect_Click()
On Error GoTo Jump
    
    STemp = "provider=sqloledb;server=" & Trim(txtserver.Text) & ";user id=" & Trim(txtuser.Text) & ";password=" & Trim(txtpass.Text) & ""

    DSN_Less_Connect STemp, SQL_Servers
    
    If ErrCount = 0 Then
    
      Screen.MousePointer = vbHourglass

      If SQL_Databases.State = 1 Then SQL_Databases.Close
      SQL_Databases.Open "sp_helpdb", cnDsn, adOpenDynamic, adLockOptimistic
      cbosqldatabase.Clear

      While Not SQL_Databases.EOF
        If SQL_Databases.Fields("name") <> "master" And SQL_Databases.Fields("name") <> "model" And _
          SQL_Databases.Fields("name") <> "msdb" Then
          cbosqldatabase.AddItem SQL_Databases.Fields("name")
        End If
      SQL_Databases.MoveNext
      Wend
    
    Screen.MousePointer = vbArrow
   
       Height = 4310
       Top = 2370
     
    End If
    
    
Exit Sub
Jump:
MsgBox Err.Description, vbCritical
Screen.MousePointer = vbArrow
End Sub

Private Sub cmdok_Click()
 On Error GoTo Jump
 
    If Trim(cbosqldatabase.Text) = "" Then Exit Sub
    
    STemp = "provider=sqloledb;server=" & Trim(txtserver) & ";user id=" & Trim(txtuser) & ";password=" & Trim(txtpass) & ";database=" & cbosqldatabase.Text
    
    Continue = True

    DSN_Less_Connect STemp, SQL_Servers
    frmmain.FillCombo
    frmmain.lstfields.ListItems.Clear
    frmmain.StatusBar1.Panels(2).Text = "Total Records : 0"
    frmmain.StatusBar1.Panels(3).Text = "Total Fields : 0"
    For i = 1 To frmmain.lstfields.ColumnHeaders.Count
      frmmain.lstfields.ColumnHeaders(i).Text = ""
    Next
    
    
    If Err.Number = 0 Then
      
      If Raiserror = False Then
            
            Database_Name = cbosqldatabase.Text
            frmmain.lbltables.Caption = "[ " & Database_Name & " : "
            frmmain.lbltables.Caption = frmmain.lbltables.Caption & IIf(Tablecount = 1, Tablecount & " Table", Tablecount & " Tables") & " ]"
            frmmain.StatusBar1.Panels(2).Text = "Total Records : 0"
            frmmain.StatusBar1.Panels(3).Text = "Total Fields : 0"
            
            frmmain.Caption = "Local Database " & Space(2) & "[ Server : SQL Server" & Space(3) & " Database : " & cbosqldatabase.Text & Space(3) & " Connection : DSN Less ]"
            
      Else
          DSNDatabase
      End If
      
    End If
    Unload Me
    
 Exit Sub
Jump:
 MsgBox Err.Description, vbCritical
End Sub

Private Sub Command1_Click()
  DSNDatabase
  Unload Me
End Sub

Private Sub Form_Load()
 Continue = False
End Sub
