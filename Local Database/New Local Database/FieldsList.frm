VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Fieldslist 
   BackColor       =   &H00D8C7BC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fields Description"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6960
   Icon            =   "FieldsList.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhide 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<< Close "
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   2655
   End
   Begin VB.ListBox lstrefrencesfields 
      BackColor       =   &H00E0E0E0&
      Height          =   3180
      ItemData        =   "FieldsList.frx":0442
      Left            =   7080
      List            =   "FieldsList.frx":0444
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.ComboBox cboprimarykeyfields 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   7080
      TabIndex        =   7
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdrelation 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relation Details >>"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FieldsList.frx":0446
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstdesc 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   7646
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Field Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Length"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Relation"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label lblfieldscount 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refrences Tables And Fields"
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
      Height          =   195
      Left            =   7080
      TabIndex        =   11
      Top             =   720
      Width           =   2460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Refrences Tables :"
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
      Height          =   195
      Left            =   7080
      TabIndex        =   10
      Top             =   4440
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Primary key field"
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
      Height          =   195
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   2010
   End
   Begin VB.Label fieldscount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   4830
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fields :"
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
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4830
      Width           =   1125
   End
   Begin VB.Label lbltablename 
      Alignment       =   2  'Center
      BackColor       =   &H00D8C7BC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   75
      Width           =   3240
   End
End
Attribute VB_Name = "Fieldslist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num%

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const LB_SETTABSTOPS = &H192
Dim Tabstops As Long

Private Sub cboprimarykeyfields_Click()
  
   'CHECK FORIEGN KEY
   Me.lstrefrencesfields.Clear

   SendMessage Me.lstrefrencesfields.hwnd, LB_SETTABSTOPS, 1, Tabstops

   Set Fk = cn.OpenSchema(adSchemaForeignKeys)
   While Not Fk.EOF
       If Trim(Me.cboprimarykeyfields.Text) = Fk.Fields("FK_COLUMN_NAME") Then
          GetKey = "Foreign Key" & " (" & Fk.Fields("PK_TABLE_NAME") & ")"
          lstrefrencesfields.AddItem Fk.Fields("FK_TABLE_NAME") & vbTab & ":    " & Fk.Fields("FK_COLUMN_NAME") & " "
       End If

   Fk.MoveNext
   Wend

   lblfieldscount.Caption = Me.lstrefrencesfields.ListCount

End Sub

Private Sub cboprimarykeyfields_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdhide_Click()
Width = 7020
Left = 2535
End Sub

Private Sub cmdrelation_Click()
Width = 10020
Left = 1025
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub Form_Load()
  Tabstops = 70
End Sub

Private Sub lstdesc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    lstdesc.SortKey = ColumnHeader.Index - 1
    
    If num = 0 Then
      lstdesc.SortOrder = lvwAscending
      num = 1
    Else
      lstdesc.SortOrder = lvwDescending
      num = 0
    End If
   
End Sub
