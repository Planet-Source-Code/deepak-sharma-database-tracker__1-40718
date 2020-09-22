VERSION 5.00
Begin VB.Form frmODBCLogon 
   BackColor       =   &H00D8C7BC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODBC Logon"
   ClientHeight    =   2295
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4470
   Icon            =   "ODBCLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStep3 
      BackColor       =   &H00D8C7BC&
      Caption         =   "Existing Connections"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdCancel 
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
         Height          =   330
         Left            =   2445
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1480
         Width           =   840
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1480
         Width           =   840
      End
      Begin VB.CheckBox chkdefault 
         BackColor       =   &H00D8C7BC&
         Caption         =   "Set As Default"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   1500
         TabIndex        =   3
         Top             =   990
         Width           =   1335
      End
      Begin VB.ComboBox cboDSNList 
         Height          =   315
         ItemData        =   "ODBCLogon.frx":0442
         Left            =   885
         List            =   "ODBCLogon.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   3120
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   240
         Picture         =   "ODBCLogon.frx":0446
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&DSN:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   0
         Top             =   520
         Width           =   390
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
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdconnect 
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   870
         Width           =   2055
      End
      Begin VB.TextBox txtuser 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   390
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         Picture         =   "ODBCLogon.frx":0A27
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User name"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmODBCLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1

Enum DSNTypes
   SQL_ServerDSN = 1
   MYSQlDSN = 2
   MSAccessDSN = 3
End Enum

Private Sub cboDSNList_Click()
  
   'Default Select The DSN
   If Trim(GetDsn) = Trim(cboDSNList.Text) Then
      chkdefault.Value = 1
   Else
      txtuser.Text = ""
      txtpass.Text = ""
      chkdefault.Value = 0
   End If
   
   'If DSN is for SQL Server then Get the authentication frame above for login
   If DatabaseType = SQL_Server_DSN Then
      frmsql.ZOrder
      txtuser.SelStart = Len(txtuser.Text)
   Else
      fraStep3.ZOrder
   End If
  
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdconnect_Click()
fraStep3.ZOrder
End Sub

Private Sub cmdok_Click()
On Error GoTo Jump

 Screen.MousePointer = vbHourglass
 
 If DatabaseType = SQL_Server_DSN Then
    Connect cboDSNList.Text, txtuser.Text, txtpass.Text
 Else
    Connect cboDSNList.Text
 End If
 
 If Raiserror = False Then
 
    frmmain.FillCombo
    frmmain.lstfields.ListItems.Clear
    frmmain.lbltables.Caption = "[ " & Database_Name & " : "
    frmmain.lbltables.Caption = frmmain.lbltables.Caption & IIf(Tablecount = 1, Tablecount & " Table", Tablecount & " Tables") & " ]"
    frmmain.StatusBar1.Panels(2).Text = "Total Records : 0"
    frmmain.StatusBar1.Panels(3).Text = "Total Fields : 0"
    
    For i = 1 To frmmain.lstfields.ColumnHeaders.Count
        frmmain.lstfields.ColumnHeaders(i).Text = ""
    Next
    
    If Trim(GetDsn) = Trim(cboDSNList.Text) Then
       
       If chkdefault.Value = 0 Then
          SetDsn ""
          SetDsnDatabase ""
          If DatabaseType = SQL_Server_DSN Then SetAuthentication ""
       Else
          If DatabaseType = SQL_Server_DSN Then SetAuthentication Trim(txtuser.Text) & "|" & Trim(txtpass.Text) & "|"
       End If
       
    Else
    
       Select Case DatabaseType
         Case MSAccess_DSN:    SetDsn (cboDSNList.Text): SetDsnDatabase ("MS Access"): SetAuthentication ""
         Case SQL_Server_DSN:  SetDsn (cboDSNList.Text): SetDsnDatabase ("SQL Server"): SetAuthentication Trim(txtuser.Text) & "|" & Trim(txtpass.Text) & "|"
         Case MYSQl:           SetDsn (cboDSNList.Text): SetDsnDatabase ("MySQL"): SetAuthentication ""
       End Select
       
    End If
       
    Select Case DatabaseType
      Case MSAccess_DSN: frmmain.Caption = "Local Database " & Space(2) & "[ Database : MS Access" & Space(3) & " DSN : " & Trim(cboDSNList.Text) & " ]"
      Case SQL_Server_DSN: frmmain.Caption = "Local Database " & Space(2) & "[ Database : SQL Server " & Space(3) & " DSN : " & Trim(cboDSNList.Text) & " ]"
      Case MYSQl: frmmain.Caption = "Local Database " & Space(2) & "[ Database : MySQL" & Space(3) & " DSN : " & Trim(cboDSNList.Text) & " ]"
    End Select
    
  End If
  
 Screen.MousePointer = vbArrow
 Unload Me
Exit Sub
Jump:
 MsgBox Err.Description, vbInformation
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    
   Select Case DatabaseType
   Case MSAccess_DSN
       
       fraStep3.Caption = "Existing Connections For (MS Access)"
       GetDSNs MSAccessDSN
         
   Case SQL_Server_DSN
       
       fraStep3.Caption = "Existing Connections For (SQL Server)"
       GetDSNs SQL_ServerDSN
       GetAuthentication_Information
       txtuser.Text = SQL_Authentication(0).UID
       txtpass.Text = SQL_Authentication(1).Pass
       
   Case MYSQl
       
       fraStep3.Caption = "Existing Connections For (MYSQL)"
       GetDSNs MYSQlDSN
       
   End Select
   
   For i = 0 To cboDSNList.ListCount - 1
   
     If UCase(Trim(GetDsn)) = UCase(Trim(cboDSNList.List(i))) Then
       cboDSNList.ListIndex = i
       Exit For
     End If
   
   Next
   
   fraStep3.ZOrder
    
End Sub

Sub GetDSNs(dsns As DSNTypes)
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         'handle to the environment

    'On Error Resume Next

    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
            
            If sDSN <> Space(iDSNLen) Then
            
              If dsns = MSAccessDSN Then
              
                    If Trim(UCase(Mid(sDRV, 1, Len(MSAccess_Tag)))) = Trim(UCase(MSAccess_Tag)) Then
                       If sDSN <> "MS Access 97 Database" And sDSN <> "MS Access Database" Then
                   
                          cboDSNList.AddItem sDSN
                        
                       End If
                        
                    End If
                    
                ElseIf dsns = MYSQlDSN Then
                  
                     If Trim(UCase(Mid(sDRV, 1, Len(MYSQL_Tag)))) = Trim(UCase(MYSQL_Tag)) Then
                       
                          cboDSNList.AddItem sDSN
                        
                     End If
                  
                ElseIf dsns = SQL_ServerDSN Then
                  
                     If Trim(UCase(Mid(sDRV, 1, Len(SQlServer_Tag)))) = Trim(UCase(SQlServer_Tag)) Then
                       
                          cboDSNList.AddItem sDSN
                        
                     End If
                  
                End If
                
            End If
        Loop
    End If
End Sub


