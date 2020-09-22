Attribute VB_Name = "Main"
'----------------------------------------------------------
' Design & Developed By   :   Deepak Sharma
'
' E-Mail                  :   deepakmailto@rediffmail.com
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

Public Const SQlServer_Tag = "SQL Server"
Public Const Oracle_Tag = "Microsoft ODBC For Oracle"
Public Const MSAccess_Tag = "Microsoft Access"
Public Const MYSQL_Tag = "MySQL ODBC"

Public cn                As ADODB.Connection
Public cnDsn             As ADODB.Connection

Public Table             As ADOX.Table
Public mCat              As New ADOX.Catalog

Public Field             As New ADODB.Recordset
Public SQL_Databases     As New ADODB.Recordset
Public Fill              As New ADODB.Recordset
Public Pk                As New ADODB.Recordset
Public Fk                As New ADODB.Recordset

Public i                 As Integer
Public Tablecount        As Integer
Public P                 As Integer

Public GetSet            As String
Dim TempConnect          As String
Public Database_Name     As String
Public STemp             As String

Public Temp_Auth         As Variant

Public Continue          As Boolean
Public Raiserror         As Boolean

Public ErrCount          As Long

Enum ErrorTypes
   LinkBreaks = 1
   NoConnect = 2
End Enum

Enum DSN_Less_Database
   MSAccesss = 1
   SQL_Servers = 2
End Enum

Type Authent
    UID As String
    Pass As String
End Type

Enum DatabasesTypes
   SQL_Server_DSN = 1
   SQL_Server = 2
   MYSQl = 3
   MSAccess_DSN = 4
   MSAccess = 5
End Enum

Public DatabaseType             As DatabasesTypes
Public SQL_Authentication(1)    As Authent

Public Sub Connect(Dsn_Name As String, Optional User, Optional Pass)
On Error GoTo Jump
 
    Set cnDsn = New ADODB.Connection
    
    If DatabaseType = SQL_Server_DSN Then
       cnDsn.Open "DSN=" & Dsn_Name & ";UId=" & Trim(User) & ";Pass=" & Trim(Pass) & ";"
    Else
       cnDsn.Open "DSN=" & Dsn_Name
    End If
  
    If Err.Number = 0 Then
   
      Database_Name = frmODBCLogon.cboDSNList.Text
      Set cn = New ADODB.Connection
      
      If DatabaseType = SQL_Server_DSN Then
         cn.Open "DSN=" & Dsn_Name & ";UId=" & Trim(User) & ";Pass=" & Trim(Pass) & ";"
      Else
         cn.Open "DSN=" & Dsn_Name
      End If
      
      cn.CursorLocation = adUseClient
      Set mCat.ActiveConnection = cn
      Raiserror = False
     
   Else
   
      Database_Name = ""
      Tablecount = 0
      Raiserror = True
      
   End If

Exit Sub
Jump:
  MsgBox Err.Description, vbCritical
  Raiserror = True
End Sub

Public Sub DSN_Less_Connect(ByVal Connect_String As String, Databases As DSN_Less_Database)
On Error GoTo Jump
  
  Set cnDsn = New ADODB.Connection
  TempConnect = ""
  
  Screen.MousePointer = vbHourglass
  
  If Databases = MSAccesss Then
       
        cnDsn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Connect_String & ";Persist Security Info=False"
        
        If Err.Number = 0 Then
        
            Set cn = New ADODB.Connection
            cn.Open cnDsn.ConnectionString
            cn.CursorLocation = adUseClient
            Set mCat.ActiveConnection = cn
            Database_Name = Connect_String
            frmmain.lbltables.Caption = "[ " & Database_Name & " : "
            frmmain.lbltables.Caption = frmmain.lbltables.Caption & IIf(Tablecount = 1, Tablecount & " Table", Tablecount & " Tables") & " ]"
            Raiserror = False
        
        Else
            
                Raiserror = True
                  
        End If
        
  ElseIf Databases = SQL_Servers Then
  
         cnDsn.Open Connect_String
         ErrCount = 0
         
      If Continue = True Then
         If Err.Number = 0 Then
        
              Set cn = New ADODB.Connection
              cn.Open cnDsn.ConnectionString
              cn.CursorLocation = adUseClient
              Set mCat.ActiveConnection = cn
              Raiserror = False

          Else
          
              Raiserror = True
              
         End If
     End If
         
  End If
   
    Screen.MousePointer = vbArrow

Exit Sub
Jump:
  MsgBox Err.Description, vbCritical
  ErrCount = Err.Number
  Raiserror = True
  Screen.MousePointer = vbArrow
End Sub

Public Function AddSlashes(StrVar As String) As String

    Dim cnt, NewStrVar: NewStrVar = ""
    StrVar = Trim(StrVar)
    For cnt = 1 To Len(StrVar)
       If Mid(StrVar, cnt, 1) = "'" Or Mid(StrVar, cnt, 1) = "\" Then
             NewStrVar = NewStrVar & "\"
       End If
       NewStrVar = NewStrVar & Mid(StrVar, cnt, 1)
    Next
    AddSlashes = NewStrVar
    
End Function

Public Sub DSNDatabase()
 Select Case Trim(GetDsnDatabase)
    Case "MS Access":    DatabaseType = MSAccess_DSN
    Case "SQL Server":   DatabaseType = SQL_Server_DSN
    Case "MySQL":        DatabaseType = MYSQl
 End Select
End Sub

Public Sub GetAuthentication_Information()
On Error Resume Next
   Temp_Auth = Split(Trim(GetAuthentication), "|")
   SQL_Authentication(0).UID = Temp_Auth(0)
   SQL_Authentication(1).Pass = Temp_Auth(1)

End Sub

'--------- [ REGISTRY SETTING ] ----------

Public Function GetDsn() As String
    GetDsn = GetSetting("DefaultSettings", "Settings", "DSN_Name")
End Function

Public Function SetDsn(sDsn_name As String)
    SaveSetting "DefaultSettings", "Settings", "DSN_Name", sDsn_name
End Function

Public Function GetDsnDatabase() As String
    GetDsnDatabase = GetSetting("DefaultSettings", "Settings", "Database_Name")
End Function

Public Function SetDsnDatabase(sDsnDatabase_name As String)
    SaveSetting "DefaultSettings", "Settings", "Database_Name", sDsnDatabase_name
End Function

Public Function GetAuthentication() As String
    GetAuthentication = GetSetting("DefaultSettings", "Settings", "User_Pass")
End Function

Public Function SetAuthentication(sUser_Pass As String)
    SaveSetting "DefaultSettings", "Settings", "User_Pass", sUser_Pass
End Function

Public Function GetLocalDatabasePath() As String
    GetLocalDatabasePath = GetSetting("DefaultSettings", "Settings", "Local_Database_Path")
End Function

Public Function SetLocalDatabasePath(sPath As String)
    SaveSetting "DefaultSettings", "Settings", "Local_Database_Path", sPath
End Function





