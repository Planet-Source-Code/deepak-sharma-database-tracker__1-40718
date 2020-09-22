VERSION 5.00
Begin VB.Form frmformat 
   BackColor       =   &H00D8C7BC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Format (VB)"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "Format(VB).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select All Formated Text"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7350
      Width           =   2295
   End
   Begin VB.TextBox txtformatstring 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   4080
      Width           =   8295
   End
   Begin VB.TextBox txttobeformat 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   8295
   End
   Begin VB.TextBox txtaftertotal 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7920
      TabIndex        =   7
      Top             =   7350
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Close"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7350
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Format"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7350
      Width           =   1335
   End
   Begin VB.TextBox txtwordsinline 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7680
      TabIndex        =   4
      Text            =   "50"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2640
      Picture         =   "Format(VB).frx":0442
      Top             =   750
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2640
      Picture         =   "Format(VB).frx":0544
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Text After Formated"
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
      TabIndex        =   9
      Top             =   3840
      Width           =   2115
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Lines Formated"
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
      Left            =   6000
      TabIndex        =   8
      Top             =   7380
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Words Per Line "
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
      Left            =   5640
      TabIndex        =   3
      Top             =   765
      Width           =   1905
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Format(VB).frx":0646
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
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9225
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Text To Be Formated "
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
      TabIndex        =   1
      Top             =   765
      Width           =   2280
   End
End
Attribute VB_Name = "frmformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function TotalLines() As Long
 
 Dim arr As Variant
 arr = Split(txtformatstring.Text & vbCrLf, vbNewLine)
 TotalLines = UBound(arr)
 
End Function

Private Sub Command1_Click()
If Trim(txttobeformat.Text) = "" Then
  MsgBox "No SQL Text Selected To Format", vbExclamation
  Exit Sub
End If
txtformatstring.Text = ""
txtformatstring = FormatSQL(txttobeformat.Text)
txtaftertotal.Text = TotalLines()
End Sub

Function FormatSQL(Query As String) As String
On Error Resume Next
Dim Actual
Dim WordsPerLine As Long

Query = Replace(Query & vbCrLf, vbCrLf, "")

Query = Replace(Query, """", "'")

WordsPerLine = Val(txtwordsinline.Text)

For i = 0 To Len(txttobeformat.Text)
  
  If Mid(Trim(Query), (i * WordsPerLine) + 1, WordsPerLine) <> "" Then
      
     Actual = Actual & """" & Mid(Trim(Query), (i * WordsPerLine) + 1, WordsPerLine) & """ & _" & vbCrLf
      
  End If
    
Next

Actual = Mid(Actual, 1, Len(Actual) - 5)    'remove the & _ from last

FormatSQL = Actual

End Function

Private Sub Command2_Click()

txtformatstring.SelStart = 0
txtformatstring.SelLength = Len(txtformatstring)
txtformatstring.SetFocus
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub

