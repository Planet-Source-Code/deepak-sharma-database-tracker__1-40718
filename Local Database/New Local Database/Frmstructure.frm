VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmstructure 
   BackColor       =   &H00D8C7BC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Structures"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "Frmstructure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8C7BC&
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   6975
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         Picture         =   "Frmstructure.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   1335
      End
      Begin VB.CommandButton cmdupdate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Update Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         Picture         =   "Frmstructure.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   1335
      End
      Begin VB.CommandButton cmdinsert 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insert Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         Picture         =   "Frmstructure.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   1335
      End
   End
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdclose 
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txtformatstring 
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
      ForeColor       =   &H00404080&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5040
      Width           =   6975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D8C7BC&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   6975
      Begin VB.CommandButton cmdnormal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Normal SQL Structure"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdforvb 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Structure For VB"
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame main_Frame 
      BackColor       =   &H00D8C7BC&
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   6975
      Begin VB.CommandButton cmdShow 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Refresh"
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
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
      Begin VB.ListBox lstfields 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         ItemData        =   "Frmstructure.frx":0820
         Left            =   2040
         List            =   "Frmstructure.frx":0822
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   720
         Width           =   3375
      End
      Begin VB.ListBox lstwherelist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         ItemData        =   "Frmstructure.frx":0824
         Left            =   2040
         List            =   "Frmstructure.frx":0826
         MultiSelect     =   1  'Simple
         TabIndex        =   11
         Top             =   2040
         Width           =   3375
      End
      Begin MSForms.ComboBox cbotable 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   3375
         VariousPropertyBits=   746604571
         ForeColor       =   8388608
         DisplayStyle    =   3
         Size            =   "5953;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblhead3 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8C7BC&
         Caption         =   "Where"
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
         Left            =   600
         TabIndex        =   10
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label lblhead1 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8C7BC&
         Caption         =   "Insert Into"
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
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblhead2 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8C7BC&
         Caption         =   "Fields"
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
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   510
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8C7BC&
      Caption         =   "Generate SQL Structure"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   2445
   End
End
Attribute VB_Name = "frmstructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL            As String
Public Tags        As String
Dim Counter        As Integer

Dim Selected       As Boolean
Dim tablefound     As Boolean

Dim Fieldslist     As Variant
Dim WhereLists     As Variant
Dim Temp           As Variant
Dim WhereList      As Variant
Dim Lists          As Variant
Dim ValuesList     As Variant

Private Sub cbotable_Click()
cmdShow_Click
End Sub

Private Sub cbotable_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
On Error GoTo Jump
If KeyCode = 13 Then

cmdShow_Click

End If

Exit Sub

Jump:
  
 If Err.Number = "-2147217865" Then
   MsgBox "No A Valid Table", vbCritical
   lstfields.Clear
   lstwherelist.Clear
 Else
   MsgBox Err.Description, vbCritical
 End If

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
  main_Frame.ZOrder
  lblhead1.Caption = "Delete From"
  lblhead2.Caption = "Where"
  lblhead3.Caption = ""
  lstfields.Height = 2600
  Tags = "delete"
  RefreshSelect
  cmdinsert.BackColor = &H8000000F
  cmdupdate.BackColor = &H8000000F
  cmddelete.BackColor = &HC0C000
End Sub

Private Sub cmdforvb_Click()

If cbotable.Text = "" Then
  MsgBox "Select Table Name", vbExclamation
  Exit Sub
End If

tablefound = False

For i = 0 To cbotable.ListCount - 1
    If Trim(cbotable.Text) = cbotable.List(i) Then
       tablefound = True
       Exit For
    End If
Next

If tablefound = False Then
  MsgBox "No A Valid Table", vbExclamation
  cbotable.Text = ""
  lstfields.Clear
  lstwherelist.Clear
  Exit Sub
End If
 
Selected = False
 
For i = 0 To lstfields.ListCount - 1

    If lstfields.Selected(i) = True Then
       
       Selected = True
       
    End If

Next
 
If Selected = False Then
  
  MsgBox "Select At Least One Field From The Field List", vbExclamation
  Exit Sub

End If
 
 txtformatstring.Text = ""
 txtformatstring.Text = FormatSQL
 
End Sub

Function FormatSQL() As String

    WhereList = ""
    WhereLists = ""
    ValuesList = ""
    Lists = ""
    Counter = 1
   
   For i = 0 To lstfields.ListCount - 1
   
     If lstfields.Selected(i) = True Then
        
       If Tags = "insert" Then
          
          ValuesList = ValuesList & lstfields.List(i) & ","
          
          If Counter < lstfields.SelCount Then
            Lists = Lists & """'""& " & lstfields.List(i) & " &""',""" & " & _" & vbCrLf
          Else
            Lists = Lists & """'""& " & lstfields.List(i) & " &""'""" & " & _" & vbCrLf
          End If
          
          Counter = Counter + 1
          
       ElseIf Tags = "update" Then
       
          If ValuesList = "" Then ValuesList = """"
          ValuesList = ValuesList & lstfields.List(i) & "='""" & " & " & lstfields.List(i) & " & " & """'"" & _ " & vbCrLf & ""","
          
       ElseIf Tags = "delete" Then
          
          If WhereLists = "" Then
            WhereLists = WhereLists & """" & lstfields.List(i) & "='""" & " & " & lstfields.List(i) & " & " & """'"" & _ " & vbCrLf & """and "
          Else
            WhereLists = WhereLists & lstfields.List(i) & "='""" & " & " & lstfields.List(i) & " & " & """'"" & _ " & vbCrLf & """and "
          End If
       
       End If
       
     End If
   
   Next
   
   'where for update only
   
   For i = 0 To lstwherelist.ListCount - 1
   
     If lstwherelist.Selected(i) = True Then

         WhereList = WhereList & lstwherelist.List(i) & "='""" & " & " & lstwherelist.List(i) & " & " & """'"" & _ " & vbCrLf & """and "
        
     End If
     
   Next
   
   Select Case Tags
   Case "insert"
       
         FormatSQL = """Insert into " & cbotable & "(" & Mid(ValuesList, 1, Len(ValuesList) - 1) & ") values(""" & " & _ " & vbCrLf & Lists & """)"""
       
   Case "update"
   
        If ValuesList <> "" And WhereList = "" Then
       
         FormatSQL = """Update " & cbotable & " Set " & """ & _" & vbCrLf & Mid(ValuesList, 1, Len(ValuesList) - 8)
         
        Else
        
         FormatSQL = """Update " & cbotable & " Set " & """ & _" & vbCrLf & Mid(ValuesList, 1, Len(ValuesList) - 1) & "Where " & Mid(WhereList, 1, Len(WhereList) - 11)
         
        End If
        
   Case "delete"
   
        FormatSQL = """Delete from " & cbotable & " Where " & """ & _" & vbCrLf & Mid(WhereLists, 1, Len(WhereLists) - 11)
   
   
   End Select

End Function

Private Function Format() As String
    
    WhereList = ""
    WhereLists = ""
    ValuesList = ""
    Lists = ""
    Temp = ""
 
 With lstfields
  
    SQL = ""
  
    For i = 0 To .ListCount - 1
     
      If .Selected(i) = True Then
            
         FieldList = FieldList & .List(i) & IIf(Tags = "insert", ",", "='' , ")
         Temp = Temp & "'',"
         WhereList = WhereList & .List(i) & "='' And "
         
      End If
      
    Next
   
    For i = 0 To lstwherelist.ListCount - 1
    
      If lstwherelist.Selected(i) = True Then
         WhereLists = WhereLists & lstwherelist.List(i) & "='' And "
      End If
      
    Next
     
    If FieldList <> "" Then FieldList = Mid(FieldList, 1, Len(FieldList) - 1)
    If Temp <> "" Then Temp = Mid(Temp, 1, Len(Temp) - 1)
    If WhereLists <> "" Then WhereLists = Mid(WhereLists, 1, Len(WhereLists) - 4)
    If WhereList <> "" Then WhereList = Mid(WhereList, 1, Len(WhereList) - 4)
   
    Select Case Tags
    
    Case "insert"
  
         Format = "Insert into " & cbotable.Text & "(" & FieldList & ") values(" & Temp & ")"
  
    Case "update"
    
         If FieldList <> "" And WhereLists = "" Then
            Format = "Update " & cbotable.Text & " Set " & Mid(FieldList, 1, Len(FieldList) - 1)
         Else
            Format = "Update " & cbotable.Text & " Set " & Mid(FieldList, 1, Len(FieldList) - 1) & " Where " & WhereLists
         End If
     
    Case "delete"
    
         If WhereList <> "" Then
            Format = "Delete From " & cbotable.Text & " Where " & WhereList
         Else
            Format = "Delete From " & cbotable.Text
         End If
  
    End Select
   
 End With
 
End Function

Private Sub cmdinsert_Click()
main_Frame.ZOrder
lblhead1.Caption = "Insert Into"
lblhead2.Caption = "Fields"
lblhead3.Caption = ""
lstfields.Height = 2600
Tags = "insert"
RefreshSelect
cmdinsert.BackColor = &HC0C000
cmdupdate.BackColor = &H8000000F
cmddelete.BackColor = &H8000000F
End Sub

Public Sub RefreshSelect()
For i = 0 To lstfields.ListCount - 1
  
  lstfields.Selected(i) = False
Next

For i = 0 To lstwherelist.ListCount - 1
  
  lstwherelist.Selected(i) = False

Next
End Sub

Private Sub cmdnormal_Click()


If cbotable.Text = "" Then
  MsgBox "Select Table Name", vbExclamation
  Exit Sub
End If

tablefound = False

For i = 0 To cbotable.ListCount - 1
    If Trim(cbotable.Text) = cbotable.List(i) Then
       tablefound = True
       Exit For
    End If
Next

If tablefound = False Then
  MsgBox "No A Valid Table", vbExclamation
  cbotable.Text = ""
  lstfields.Clear
  lstwherelist.Clear
  Exit Sub
End If

Selected = False
 
For i = 0 To lstfields.ListCount - 1

    If lstfields.Selected(i) = True Then
       
       Selected = True
       
    End If

Next
 
If Selected = False Then
  
  MsgBox "Select At Least One Field From The Field List", vbExclamation
  Exit Sub

End If


  txtformatstring.Text = ""
  txtformatstring.Text = Format
    
End Sub

Private Sub cmdShow_Click()
FillList lstfields
FillList lstwherelist
End Sub

Private Sub cmdupdate_Click()
main_Frame.ZOrder
lblhead1.Caption = "Update Table"
lblhead2.Caption = "Set (Fields)"
lblhead3.Caption = "Where"
lstfields.Height = 1300
Tags = "update"
RefreshSelect
cmdinsert.BackColor = &H8000000F
cmdupdate.BackColor = &HC0C000
cmddelete.BackColor = &H8000000F
End Sub

Public Sub FillCombo()
On Error GoTo Jump

 cbotable.Clear
 cbotable.Text = ""
 
    For Each Table In mCat.Tables
    
     If Table.Type = "TABLE" Then
     
       cbotable.AddItem Table.Name
     
     End If
    
    Next
 
Exit Sub
Jump:
   MsgBox Err.Description, vbCritical
End Sub

Public Sub FillList(lst As ListBox)
On Error GoTo Jump

 If cbotable.Text <> "" Then

     lst.Clear
     If Field.State = 1 Then Field.Close
     Field.Open "select * from " & Trim(cbotable.Text), cn, adOpenDynamic, adLockOptimistic
      
  Screen.MousePointer = vbHourglass
      
     For i = 0 To Field.Fields.Count - 1
    
        lst.AddItem Field.Fields(i).Name
    
     Next
     
   Screen.MousePointer = vbNormal
 
 End If
 
Exit Sub
Jump:

     MsgBox Err.Description, vbCritical

End Sub

Private Sub Command2_Click()
    txtformatstring.SelStart = 0
    txtformatstring.SelLength = Len(txtformatstring)
    txtformatstring.SetFocus
End Sub
