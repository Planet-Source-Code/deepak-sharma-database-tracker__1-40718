VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8C7BC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Local Database"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10155
   Icon            =   "local.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8580
   ScaleMode       =   0  'User
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6915
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Picture         =   "local.frx":0442
            Text            =   "Done    "
            TextSave        =   "Done    "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "local.frx":0894
            Text            =   "Total Records : 0"
            TextSave        =   "Total Records : 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "local.frx":0CE6
            Text            =   "Total Fields : 0"
            TextSave        =   "Total Fields : 0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   50
      TabIndex        =   2
      Top             =   120
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   11880
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   128
      TabCaption(0)   =   "Records"
      TabPicture(0)   =   "local.frx":0E40
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MenuImg"
      Tab(0).Control(1)=   "chkmulti"
      Tab(0).Control(2)=   "cd"
      Tab(0).Control(3)=   "ImageList1"
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(5)=   "Check1"
      Tab(0).Control(6)=   "Timer1"
      Tab(0).Control(7)=   "lstfields"
      Tab(0).Control(8)=   "ctrl"
      Tab(0).Control(9)=   "Image3"
      Tab(0).Control(10)=   "Image2"
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(12)=   "cbotables"
      Tab(0).Control(13)=   "lblctrlkeys"
      Tab(0).Control(14)=   "lblcolumnheads"
      Tab(0).Control(15)=   "cmddescriptions"
      Tab(0).Control(16)=   "Image1"
      Tab(0).Control(17)=   "lbltables"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Query"
      TabPicture(1)   =   "local.frx":0E5C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstresult"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txterrors"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdbrowse"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdrun"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdnew"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdbatch"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdzoom"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdformatvb"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdstructure"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdjoins"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtquery"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin MSComctlLib.ImageList MenuImg 
         Left            =   -69600
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":0E78
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":11CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":151C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":182E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtquery 
         Height          =   3255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   5741
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         BulletIndent    =   200
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"local.frx":1CB8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdjoins 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete/Drop"
         Height          =   615
         Left            =   7800
         Picture         =   "local.frx":1D7B
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   50
         Width           =   1935
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lbltotalrecords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   29
            Top             =   480
            Width           =   90
         End
         Begin VB.Label lbltotalfields 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   28
            Top             =   120
            Width           =   90
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Field        :"
            ForeColor       =   &H0080FF80&
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Records  :"
            ForeColor       =   &H0080FF80&
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdstructure 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Structures "
         Height          =   615
         Left            =   9000
         Picture         =   "local.frx":1EC5
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdformatvb 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Format (VB)"
         Height          =   615
         Left            =   6600
         Picture         =   "local.frx":200F
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdzoom 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Zoom"
         Height          =   615
         Left            =   4680
         Picture         =   "local.frx":2111
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdbatch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Batch Fire"
         Height          =   615
         Left            =   5520
         Picture         =   "local.frx":2643
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox chkmulti 
         Caption         =   "Checkboxes"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   18
         Top             =   520
         Width           =   1575
      End
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         Height          =   615
         Left            =   3000
         Picture         =   "local.frx":278D
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   -73200
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdrun 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Run (F5)"
         Height          =   615
         Left            =   2160
         Picture         =   "local.frx":288F
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Browse"
         Height          =   615
         Left            =   3840
         Picture         =   "local.frx":28D5
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -72000
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":29D7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":2CF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":3143
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "local.frx":3595
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Referesh"
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
         Left            =   -67970
         Picture         =   "local.frx":39E7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   220
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Referesh"
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
         Height          =   255
         Left            =   -74880
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   -74160
         Top             =   1320
      End
      Begin MSComctlLib.ListView lstfields 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   1
         Top             =   1200
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   9128
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
         ColHdrIcons     =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txterrors 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   4320
         Width           =   9790
      End
      Begin MSComctlLib.ListView lstresult 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   4320
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   3625
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ColHdrIcons     =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Image ctrl 
         Height          =   240
         Left            =   -68400
         Picture         =   "local.frx":3B31
         Stretch         =   -1  'True
         Top             =   930
         Width           =   315
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   -71400
         Picture         =   "local.frx":3F73
         Top             =   930
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   -74880
         Picture         =   "local.frx":4075
         Top             =   930
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type Few Words And Press Enter"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -71040
         TabIndex        =   20
         Top             =   960
         Width           =   2400
      End
      Begin MSForms.ComboBox cbotables 
         Height          =   375
         Left            =   -72480
         TabIndex        =   0
         Top             =   405
         Width           =   4335
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   8388608
         DisplayStyle    =   3
         Size            =   "7646;661"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   2
         BorderColor     =   16512
         SpecialEffect   =   6
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblctrlkeys 
         AutoSize        =   -1  'True
         Caption         =   "Hold CTRL Key And Click For MultiSelect"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -68040
         TabIndex        =   19
         Top             =   960
         Width           =   2925
      End
      Begin VB.Label lblcolumnheads 
         AutoSize        =   -1  'True
         Caption         =   "Click On The Columns Head For Sorting"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74520
         TabIndex        =   17
         Top             =   960
         Width           =   2820
      End
      Begin MSForms.CommandButton cmddescriptions 
         Height          =   630
         Left            =   -66840
         TabIndex        =   16
         Top             =   240
         Width           =   1695
         BackColor       =   14737632
         VariousPropertyBits=   25
         Caption         =   "Fields Description"
         Size            =   "2990;1111"
         Picture         =   "local.frx":4177
         FontEffects     =   1073750017
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -73080
         Picture         =   "local.frx":42D1
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "( Select specific query in the multiple queries by selecting it )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5520
         TabIndex        =   12
         Top             =   4080
         Width           =   4365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Result"
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
         Left            =   120
         TabIndex        =   8
         Top             =   4080
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type your query and press Run or F5 button."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1680
         TabIndex        =   6
         Top             =   4080
         Width           =   3270
      End
      Begin VB.Label lbltables 
         Alignment       =   2  'Center
         Caption         =   "Tables"
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
         Left            =   -72495
         TabIndex        =   5
         Top             =   120
         Width           =   4245
      End
   End
   Begin VB.Menu mnuconnect 
      Caption         =   "Connect"
      Begin VB.Menu mnuaccess 
         Caption         =   "MS Access "
      End
      Begin VB.Menu mnuSqlserver 
         Caption         =   "SQL Server"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuodbc 
         Caption         =   "ODBC (DSN)"
         Begin VB.Menu mnuaccessdsn 
            Caption         =   "MS Access"
         End
         Begin VB.Menu mnuSqlserverdsn 
            Caption         =   "SQL Server"
         End
         Begin VB.Menu mnumysql 
            Caption         =   "MySQL"
         End
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuruntime 
         Caption         =   ""
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Design & Developed By   :   Deepak Sharma
'
' E-Mail                  :   deepakmailto@rediffmail.com
'
'
' For Any Queries And Suggestions Please Write To Me At My
' E-Mail Address.
'
' If You Like This Code Then Please Vote Me And Post Yours
' Messages So That I Can Judge My Programming Skills.
'
'                                                Thanks
'                                           (Deepak Sharma)
'----------------------------------------------------------

Dim RunQuery            As New ADODB.Recordset

Dim X                   As ListItem

Dim num                 As Integer
Dim i                   As Integer
Dim j                   As Integer

Dim SQL                 As String
Public Temp             As String
Dim FieldTypes          As String
Dim TempTable           As String
Dim Tablename           As String
Dim Filename            As String
Dim TempStore           As String
Dim Current_Table       As String

Dim tablefound          As Boolean
Dim Fills               As Boolean
Dim SQLQUERY            As Boolean

Private Type FieldList
 
  Field_Name            As String
  Field_Type            As String
  Field_Length          As String

End Type

Dim StoreValues()       As FieldList

Public Sub FillGrid()
On Error GoTo Jump

   If Me.cbotables.Text <> "" Then
       
       Erase StoreValues
       Me.cmddescriptions.Enabled = True
       Screen.MousePointer = vbHourglass
       StatusBar1.Panels(1).Picture = ImageList1.ListImages(3).Picture
       StatusBar1.Panels(1).Text = "Wait ! Searching Records..."
       StatusBar1.Panels(2).Text = "Total Records : 0"
       StatusBar1.Panels(3).Text = "Total Fields : 0"
   
       Me.lstfields.ListItems.Clear
       Me.lstfields.ColumnHeaders.Clear
       
       Tablename = ""
       FieldTypes = ""
       Current_Table = Me.cbotables.Text
   
        If Fill.State = 1 Then Fill.Close
        If Field.State = 1 Then Field.Close
        
        If DatabaseType = MYSQl Then
        
          Tablename = "select * from " & Trim(Me.cbotables.Text)
        
        Else
        
          Tablename = "select * from [" & (Me.cbotables.Text) & "]"
        
        End If
        
        
        Fill.Open Tablename, cn, adOpenDynamic, adLockOptimistic
        Field.Open Tablename, cn, adOpenDynamic, adLockOptimistic
        
        ReDim StoreValues(Field.Fields.Count - 1)   'store the field names
   
        For i = 0 To Field.Fields.Count - 1

            With Me.lstfields
   
                .ColumnHeaders.Add , , Field.Fields(i).Name, 1800
                .HideSelection = True
                 
                 StoreValues(i).Field_Name = Field.Fields(i).Name
                 StoreValues(i).Field_Type = cType(Field.Fields(i).Type)
                 StoreValues(i).Field_Length = Field.Fields(i).DefinedSize
                 
            End With
   
        Next
        
        j = 0
        
        While Not Fill.EOF
    
            Set X = lstfields.ListItems.Add(, , Fill.Fields(0) & "")
    
            For i = 1 To Fill.Fields.Count - 1
    
                With Me.lstfields
      
                    X.SubItems(i) = Fill.Fields(i) & ""
       
                End With
                
            Next
    
            j = j + 1
            Fill.MoveNext
         Wend
         
         Screen.MousePointer = vbArrow
         Me.StatusBar1.Panels(1).Text = "Done"
         StatusBar1.Panels(1).Picture = ImageList1.ListImages(4).Picture
         Me.StatusBar1.Panels(2).Text = "Total Records : " & Fill.RecordCount
         Me.StatusBar1.Panels(3).Text = "Total Fields : " & Field.Fields.Count
         Me.cbotables.SetFocus
         
   End If
 
Exit Sub
Jump:

     StatusBar1.Panels(1).Text = "Done"
     StatusBar1.Panels(1).Picture = ImageList1.ListImages(4).Picture
     Me.cmddescriptions.Enabled = False
     Screen.MousePointer = vbArrow
     MsgBox Err.Description, vbCritical
     
End Sub

Private Sub cbotables_Click()
On Error GoTo Jump

TempTable = Trim(cbotables.Text)
If Fills = True Then FillGrid
Fills = False

Exit Sub
Jump:
  
     MsgBox Err.Description, vbCritical
  
End Sub

Private Sub cbotables_DropButtonClick()
 
Fills = True

End Sub

Private Sub cbotables_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

 If KeyCode = 13 Then FillGrid

End Sub

Private Sub Check1_Click()
    
    If Me.Check1.Value = 1 Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
    
End Sub

Private Sub chkmulti_Click()

If chkmulti.Value = 1 Then

   Me.lstfields.Checkboxes = True
   Me.lblctrlkeys.Visible = False
   ctrl.Visible = False
   FillGrid
   'cbotables_Click

ElseIf chkmulti.Value = 0 Then

  Me.lstfields.Checkboxes = False
  Me.lblctrlkeys.Visible = True
  ctrl.Visible = True

End If

End Sub

Private Sub cmdbatch_Click()
On Error GoTo Jump

Dim BatchQuery, TempStore As Variant
Dim Counter As Integer

If Trim(txtquery.Text) <> "" Then

  Counter = 0
  lbltotalfields.Caption = "0"
  lbltotalrecords.Caption = "0"
  txtquery.Height = 3255
  
  TempStore = ""
  TempStore = txtquery.Text
  TempStore = Replace(TempStore, vbCrLf, "|")
  BatchQuery = Split(TempStore, "|")
  
  For i = 0 To UBound(BatchQuery)
   
   If BatchQuery(i) <> "" Then
      
      TempStore = BatchQuery(i) 'store for error
      cn.Execute BatchQuery(i)
      Counter = Counter + 1
      txterrors.Visible = True
      lstresult.Visible = True
      txterrors.Text = Counter & " Row(s) Affected "

   End If
   
  Next
  
     Temp = cbotables.Text
     FillCombo
     cbotables.Text = Temp
     FillGrid
  
  Erase BatchQuery

End If

Exit Sub
Jump:
     
  If Err.Number <> 0 Then
     
     txterrors.Visible = True
     txterrors.ZOrder
     'MsgBox ExtractErrors(Err.Description)
     txterrors.Text = Counter & " Row(s) Affected " & vbCrLf & vbCrLf & "Warning :  Above Selected Query Has Some Syntax Problem Check The Error." & vbCrLf & vbCrLf & "Error :  " & Err.Description
     
     Temp = cbotables.Text
     FillCombo
     cbotables.Text = Temp
     FillGrid
     
     For i = 1 To Len(txtquery.Text)
       
       If Trim(TempStore) = Trim(Mid(txtquery.Text, i, Len(TempStore))) Then
       
        txtquery.SelStart = i - 1
        txtquery.SelLength = Len(TempStore)
        txtquery.SelColor = vbRed
        txtquery.SetFocus
        Exit For
       
       End If
       
     Next
     
     txtquery.Height = 3255
     Me.txtquery.SetFocus
     lbltotalfields.Caption = "0"
     lbltotalrecords.Caption = "0"
     
  
    Exit Sub
   End If
   
End Sub

Private Sub cmdbrowse_Click()
 
On Error Resume Next
    TempStore = ""
    With cd
      .DialogTitle = "Select SQL File"
      .Filter = "All SQL File|*.sql;*.txt"
      .ShowOpen
      If Me.cd.Filename <> "" Then
 
         Me.txtquery.LoadFile cd.Filename
         TempStore = Me.txtquery.Text
         Me.txtquery.Text = ""
         Me.txtquery.SelColor = vbBlue
         Me.txtquery.SelText = TempStore

      End If
  
    End With
    
End Sub

Private Function GetKey(j As Integer) As String
 
 'CHECK PRIMARY KEY
  Set Pk = cn.OpenSchema(adSchemaPrimaryKeys)
    While Not Pk.EOF
     If Trim(Me.cbotables.Text) = Pk.Fields("TABLE_NAME") Then
       If StoreValues(j).Field_Name = Pk.Fields("COLUMN_NAME") Then
          GetKey = "Primary Key"
          Fieldslist.cboprimarykeyfields.AddItem Pk.Fields("COLUMN_NAME")
       End If
     End If
  Pk.MoveNext
  Wend

  'CHECK FORIEGN KEY
   Set Fk = cn.OpenSchema(adSchemaForeignKeys)
   While Not Fk.EOF
     If Trim(Me.cbotables.Text) = Fk.Fields("FK_TABLE_NAME") Then
       If StoreValues(j).Field_Name = Fk.Fields("FK_COLUMN_NAME") Then
          GetKey = "Foreign Key" & " (" & Fk.Fields("PK_TABLE_NAME") & ")"
       End If
     End If
   Fk.MoveNext
   Wend

End Function

Private Sub cmddescriptions_Click()
On Error GoTo Jump

  tablefound = False

  For i = 0 To cbotables.ListCount - 1
    If Trim(cbotables.Text) = cbotables.List(i) Then
       tablefound = True
       Exit For
    End If
  Next

  If tablefound = False Then
    MsgBox "Cannot show the fields description" & vbCrLf & "    Table name does not exist. ", vbCritical
    Me.cbotables.SetFocus
    Exit Sub
  End If

  Fieldslist.lstdesc.ListItems.Clear
  Fieldslist.cboprimarykeyfields.Clear
  Fieldslist.lstrefrencesfields.Clear
  
  For j = 0 To UBound(StoreValues())
    
     Set X = Fieldslist.lstdesc.ListItems.Add(, , StoreValues(j).Field_Name, 1, 1)
     
     X.SubItems(1) = StoreValues(j).Field_Type
     X.SubItems(2) = StoreValues(j).Field_Length
     
     If DatabaseType = MSAccess Or DatabaseType = SQL_Server Then
        X.SubItems(3) = GetKey(j)
     End If
         
  Next
  Fieldslist.Form_Load
  Fieldslist.lbltablename.Caption = UCase(Me.cbotables.Text)
  Fieldslist.fieldscount.Caption = Fieldslist.lstdesc.ListItems.Count
  
  Fieldslist.Show vbModal
  
Exit Sub
Jump:
   
     MsgBox Err.Description, vbCritical
  
End Sub

Private Sub cmdformatvb_Click()

 frmformat.txttobeformat.Text = txtquery.SelText
 frmformat.txtwordsinline.Text = "50"
 frmformat.Show vbModal

End Sub

Private Sub cmdjoins_Click()
Load deletedrop
deletedrop.Show vbModal
End Sub

Private Sub cmdnew_Click()
 Me.txtquery.Text = ""
 Me.txterrors.Text = ""
 lbltotalfields.Caption = "0"
 lbltotalrecords.Caption = "0"
 txtquery.SetFocus
End Sub

Private Sub cmdrun_Click()

  If Me.txtquery.Text <> "" Then

       SQL = ""
       Me.txterrors.Text = ""
       Me.txterrors.Visible = True
       Me.lstresult.Visible = True
       
       If Me.txtquery.SelText = "" Then
          SQL = Trim(Me.txtquery.Text)
       Else
          SQL = Trim(Me.txtquery.SelText)
       End If
    
       SQLQUERY = IIf(LCase(Left(SQL, 6)) <> LCase("select"), False, True)
          
       On Error GoTo Jump
        
       If RunQuery.State = 1 Then RunQuery.Close
       RunQuery.Open SQL, cn, adOpenDynamic, adLockOptimistic
       
       If SQLQUERY = True Then
       
          Me.lstresult.ZOrder
          Me.lstresult.ListItems.Clear
          Me.lstresult.ColumnHeaders.Clear
        
          For i = 0 To RunQuery.Fields.Count - 1
    
             With Me.lstresult
       
               .ColumnHeaders.Add , , RunQuery.Fields(i).Name, 1800
               .HideSelection = True
       
             End With
       
          Next
      
          If RunQuery.RecordCount > 0 Then RunQuery.MoveFirst
      
          j = 0
          lbltotalfields.Caption = RunQuery.Fields.Count
         If RunQuery.RecordCount > 0 Then
         
            lbltotalfields.Caption = RunQuery.Fields.Count
            lbltotalrecords.Caption = RunQuery.RecordCount
            txtquery.Height = 3255
      
          While Not RunQuery.EOF
        
             Set X = lstresult.ListItems.Add(, , RunQuery.Fields(0) & " ")
        
             For i = 1 To RunQuery.Fields.Count - 1
        
                With Me.lstresult
          
                  X.SubItems(i) = RunQuery.Fields(i) & " "
           
                End With
        
             Next
        
             j = j + 1
             RunQuery.MoveNext
          Wend
          
        Else
          
          lbltotalrecords.Caption = "0"
        
        End If
        
        ElseIf SQLQUERY = False Then
        
           txterrors.ZOrder
         
           txterrors.Text = "Command Completed Successfully."
           
           Temp = cbotables.Text
           FillCombo
           cbotables.Text = Temp
           FillGrid
         
        End If
        
      End If
      
        
       txtquery.Height = 3255
   Exit Sub
Jump:
  
     txterrors.ZOrder
     txterrors.Text = Err.Description
     txtquery.Height = 3255
     txtquery.SetFocus
     RunQuery.CancelUpdate
     lbltotalfields.Caption = "0"
     lbltotalrecords.Caption = "0"
     Exit Sub
   
End Sub

Private Sub cmdstructure_Click()

    Unload frmstructure
    frmstructure.FillCombo
    frmstructure.lstfields.Height = 2600
    frmstructure.lblhead3.Caption = ""
    frmstructure.lstfields.Clear
    frmstructure.lstwherelist.Clear
    frmstructure.txtformatstring.Text = ""
    frmstructure.cmdinsert.BackColor = &HC0C000
    frmstructure.cmddelete.BackColor = -2147483633
    frmstructure.cmdupdate.BackColor = -2147483633
    frmstructure.Tags = "insert"
    Load frmstructure
    frmstructure.Show vbModal

End Sub

Private Sub cmdzoom_Click()
    txtquery.Height = 5535
    Me.txterrors.Visible = False
    Me.lstresult.Visible = False
    Me.txtquery.SetFocus
End Sub

Private Sub Command1_Click()
On Error Resume Next

    If Current_Table <> "" Then
        If cbotables.ListCount = 0 Then Current_Table = ""
        Me.cbotables.Text = Current_Table
        FillGrid
    End If

End Sub

Private Sub Command5_Click()
txtquery.Height = txtquery.Height - 100
MsgBox txtquery.Height
End Sub

Private Sub Form_Load()
On Error GoTo Jump

         If Trim(GetDsn) <> "" Then
         
            Set cn = New ADODB.Connection
            
            DSNDatabase
         
            If DatabaseType = SQL_Server_DSN Then
                  
                  GetAuthentication_Information
                   
                   Connect Trim(GetDsn), Trim(SQL_Authentication(0).UID), Trim(SQL_Authentication(1).Pass)
                   
            Else
                   Connect Trim(GetDsn)
            End If
            
            If Raiserror = False Then
                
                FillCombo
                lstfields.ListItems.Clear
                frmmain.lbltables.Caption = "[ " & Trim(GetDsn) & " : "
                frmmain.lbltables.Caption = frmmain.lbltables.Caption & IIf(Tablecount = 1, Tablecount & " Table", Tablecount & " Tables") & " ]"
                
                Caption = "Local Database " & Space(2) & "[ Database : " & Trim(GetDsnDatabase) & Space(3) & " DSN : " & Trim(GetDsn) & " ]"
                
            End If
             
         Else
         
             Caption = "Local Database "
         
         End If
         
          If Trim(GetLocalDatabasePath) <> "" Then
             
             mnusep2.Visible = True
             mnuruntime.Visible = True
             mnuruntime.Caption = Trim(GetLocalDatabasePath)
             
          Else
             
              mnusep2.Visible = False
              mnuruntime.Visible = False
              mnuruntime.Caption = ""
          
          End If
         
          txtquery.Text = ""
          SSTab1.Tab = 0
          txtquery.SelColor = vbBlue
          StatusBar1.Panels(2).Text = "Total Records : 0"
          StatusBar1.Panels(3).Text = "Total Fields : 0"
          j = 0
          num = 0
          
Exit Sub
Jump:
  MsgBox Err.Description, vbCritical
End Sub

Public Function cType(ByVal Value As ADOX.DataTypeEnum) As String
  Select Case Value
    Case adTinyInt: cType = "TinyInt"
    Case adSmallInt: cType = "SmallInt"
    Case adInteger: cType = "Number"
    Case adBigInt: cType = "BigInt"
    Case adUnsignedTinyInt: cType = "UnsignedTinyInt"
    Case adUnsignedSmallInt: cType = "UnsignedSmallInt"
    Case adUnsignedInt: cType = "UnsignedInt"
    Case adUnsignedBigInt: cType = "UnsignedBigInt"
    Case adSingle: cType = "Single"
    Case adDouble: cType = "Double"
    Case adCurrency: cType = "Currency"
    Case adDecimal: cType = "Decimal"
    Case adNumeric: cType = "Numeric"
    Case adBoolean: cType = "Boolean"
    Case adUserDefined: cType = "UserDefined"
    Case adVariant: cType = "Variant"
    Case adGUID: cType = "GUID"
    Case adDate: cType = "Date/Time"
    Case adDBDate: cType = "Date/Time"
    Case adDBTime: cType = "Date/Time"
    Case adDBTimeStamp: cType = "Date/Time"
    Case adBSTR: cType = "BSTR"
    Case adChar: cType = "Text"
    Case adVarChar: cType = "Text"
    Case adLongVarChar: cType = "Text"
    Case adWChar: cType = "Text"
    Case adVarWChar: cType = "Text"
    Case adLongVarWChar: cType = "Memo"
    Case adBinary: cType = "adBinary"
    Case adVarBinary: cType = "adVarBinary"
    Case adLongVarBinary: cType = "OLE Object"
    Case Else: cType = Value
  End Select
End Function

Public Sub FillCombo()
On Error GoTo Jump

 cbotables.Clear
 cbotables.Text = ""
 Tablecount = 0
 
    For Each Table In mCat.Tables
    
     If Table.Type = "TABLE" Then
     
       cbotables.AddItem Table.Name
       Tablecount = Tablecount + 1
     
     End If
    
    Next
 
Exit Sub
Jump:

     MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 End
End Sub

Private Sub lstfields_Click()

 If Me.chkmulti = 1 Then
    For i = 1 To lstfields.ListItems.Count
        If lstfields.ListItems.Item(i).Checked = True Then
            lstfields.ListItems.Item(i).Selected = True
        Else
            lstfields.ListItems.Item(i).Selected = False
        End If
    Next
 End If

End Sub

Private Sub lstfields_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    lstfields.SortKey = ColumnHeader.Index - 1
    
    If num = 0 Then
      Me.lstfields.SortOrder = lvwAscending
      num = 1
    Else
      Me.lstfields.SortOrder = lvwDescending
      num = 0
    End If
   
End Sub

Private Sub lstresult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Me.lstresult.SortKey = ColumnHeader.Index - 1
    
    If num = 0 Then
      Me.lstresult.SortOrder = lvwAscending
      num = 1
    Else
      Me.lstresult.SortOrder = lvwDescending
      num = 0
    End If
End Sub

Private Sub mnuabout_Click()
       
  MsgBox "All rights ©reserved to Deepak Sharma" & vbCrLf & vbCrLf _
  + Space(8) & "deepakmailto@rediffmail.com"
  
End Sub

Private Sub mnuaccess_Click()

DatabaseType = MSAccess
With cd
 .DialogTitle = "Select Database"
 .Filter = "(*.MDB)|*.mdb"
 .ShowOpen
 
 If .FileTitle <> "" Then
    DSN_Less_Connect .Filename, MSAccesss
  
  If Raiserror = False Then
  
    Database_Name = .FileTitle
    FillCombo
    lstfields.ListItems.Clear
    frmmain.lbltables.Caption = "[ " & Database_Name & " : "
    frmmain.lbltables.Caption = frmmain.lbltables.Caption & IIf(Tablecount = 1, Tablecount & " Table", Tablecount & " Tables") & " ]"
    StatusBar1.Panels(2).Text = "Total Records : 0"
    StatusBar1.Panels(3).Text = "Total Fields : 0"
    For i = 1 To frmmain.lstfields.ColumnHeaders.Count
       frmmain.lstfields.ColumnHeaders(i).Text = ""
    Next
    
    Caption = "Local Database " & Space(2) & "[ " & .Filename & " ]"
    mnusep2.Visible = True
    mnuruntime.Visible = True
    mnuruntime.Caption = .Filename
    SetLocalDatabasePath .Filename
    
  End If
    
 Else
  
    DSNDatabase
    
 End If
   
End With
End Sub

Private Sub mnuaccessdsn_Click()
  DatabaseType = MSAccess_DSN
  DoEvents
  frmODBCLogon.Show 1
End Sub

Private Sub mnumysql_Click()
  DatabaseType = MYSQl
  frmODBCLogon.Show 1
End Sub

Private Sub mnuoracle_Click()
  DatabaseType = Oracle
  frmODBCLogon.Show 1
End Sub

Private Sub mnuruntime_Click()
    
    If Dir(GetLocalDatabasePath, vbNormal) = "" Then
       
        MsgBox "Cannot Find The Database File " & vbCrLf & GetLocalDatabasePath, vbCritical
        Exit Sub
    
    Else
    
        DatabaseType = MSAccess
    
        DSN_Less_Connect Trim(GetLocalDatabasePath), MSAccesss
        
        If Raiserror = False Then
        
            FillCombo
            Database_Name = Mid(Trim(GetLocalDatabasePath), InStrRev(Trim(GetLocalDatabasePath), "\") + 1)
            lstfields.ListItems.Clear
            frmmain.lbltables.Caption = "[ " & Database_Name & " : "
            frmmain.lbltables.Caption = frmmain.lbltables.Caption & IIf(Tablecount = 1, Tablecount & " Table", Tablecount & " Tables") & " ]"
            StatusBar1.Panels(2).Text = "Total Records : 0"
            StatusBar1.Panels(3).Text = "Total Fields : 0"
            For i = 1 To frmmain.lstfields.ColumnHeaders.Count
               frmmain.lstfields.ColumnHeaders(i).Text = ""
            Next
            
            Caption = "Local Database " & Space(2) & "[ " & GetLocalDatabasePath & " ]"
          
        End If
        
    End If
End Sub

Private Sub mnuSqlserver_Click()
  DatabaseType = SQL_Server
  frmSQLSERVER.Show 1
End Sub

Private Sub mnuSqlserverdsn_Click()
  DatabaseType = SQL_Server_DSN
  frmODBCLogon.Show 1
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  
  Select Case PreviousTab
  
  Case 0 'table tab
     
     If Me.Check1.Value = 1 Then Timer1.Enabled = False
     Me.txtquery.SetFocus
     num = 0
     
       
  Case 1 'query tab
  
     Temp = cbotables.Text
     Check1_Click
     FillCombo
     FillGrid
     cbotables.Text = Temp
     num = 0
     
  End Select
  
End Sub

Private Sub Timer1_Timer()
    
    Command1_Click
    
End Sub

Private Sub txtquery_Change()
   Me.txtquery.SelColor = vbBlue
End Sub

Private Sub txtquery_GotFocus()
   Me.txtquery.SelColor = vbBlue
End Sub

Private Sub txtquery_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.txtquery.SelColor = vbBlue
   If KeyCode = vbKeyF5 Then
     cmdrun_Click
   End If
End Sub

Private Sub txtquery_KeyPress(KeyAscii As Integer)
   Me.txtquery.SelColor = vbBlue
End Sub

