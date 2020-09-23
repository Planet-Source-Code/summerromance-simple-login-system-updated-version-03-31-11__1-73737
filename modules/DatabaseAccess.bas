Attribute VB_Name = "DatabaseAccess"
Public loginfail As Integer
Public db As ADODB.Connection
Public loginSucceed As Boolean, validuser(2) As Boolean, userlocked As Boolean
Public failname As String, loggedname As String, temp As String, acctid As Long
Public actL As String, AcctInfo As String

Public Sub dbconnect()
On Error GoTo errs
Screen.MousePointer = vbHourglass

Set db = New Connection
db.CursorLocation = adUseClient
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=Database\Data.mdb;" & _
        "Persist Security Info=True;" & _
        "Jet OLEDB:Database Password=TechServeEra;" & _
        "Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=True"

Screen.MousePointer = vbDefault
Exit Sub

errs:
Screen.MousePointer = vbDefault
MsgBox "Failed to Connect to database   ", vbExclamation, ""
MsgBox Err.Description
End

End Sub

Public Sub authenticate(ByVal userid As String, ByRef userpass As String)
Dim users As ADODB.Recordset
Dim check(2) As String
dbconnect

Set users = New ADODB.Recordset
   With users
        .Open ("select IDnumber,userID, userPassword from T_users_acct where userID='" & userid & "'"), db, adOpenForwardOnly, adLockReadOnly
         If .RecordCount = 1 Then
            check(0) = .Fields("userID")
            If userid = check(0) Then
               validuser(0) = True
               acctid = .Fields("IDnumber")
            Else
               validuser(0) = False
            End If
         Else
            validuser(0) = False
         End If
         .Close
   End With
       
If validuser(0) = True Then
Set users = New ADODB.Recordset
    With users
         .Open ("select acctStatus from T_users_acct where IDnumber = " & CDbl(acctid) & ""), db, adOpenForwardOnly, adLockReadOnly
         If .Fields("acctStatus") = "Locked" Then
            MsgBox "Account is locked please contact your system administrator     ", vbInformation, ""
            userlocked = True
            Exit Sub
         Else
            validuser(1) = True
         End If
         .Close
    End With
    
If validuser(1) = True Then
   EncryptDecrypt (userpass)
   userpass = temp

Set users = New ADODB.Recordset
    With users
         .Open ("select userID, userPassword, IDnumber from T_users_acct where userPassword='" & userpass & "' and userID='" & check(0) & "'"), db, adOpenForwardOnly, adLockReadOnly
         If .RecordCount = 1 Then
             check(1) = .Fields("userPassword")
             If userpass = check(1) Then
                acctid = .Fields("IDnumber")
                .Close
                Set users = Nothing
                  Set users = New ADODB.Recordset
                      .Open ("select uLevel from T_users_acct where IDnumber = " & CDbl(acctid) & " "), db, adOpenForwardOnly, adLockReadOnly
                      loggedname = .Fields("uLevel")
                      loginSucceed = True
                      validuser(2) = True
                      Call UserLogs(acctid, "Logged In")
                  .Close
                  Set users = Nothing
             Else
                validuser(2) = False
             End If
         Else
            validuser(2) = False
         End If
    End With
End If
    
End If

If validuser(0) = False Then
   MsgBox "Invalid Username    ", vbInformation
ElseIf validuser(2) = False Then
   MsgBox "Invalid Password    ", vbInformation
   loginSucceed = False
End If
Set db = Nothing

End Sub

Public Sub EncryptDecrypt(stval As String)
Dim i As Integer
Dim intLocation As Integer
Dim Code As String
Code = "1234567890" 'This is key for encrypting/decrypting
  temp = ""
  For i% = 1 To Len(stval)
      intLocation% = (i% Mod Len(Code)) + 1
      'Use XOR logic combination for encrypting/decrypting
      temp = temp + Chr$(Asc(Mid$(stval, i%, 1)) Xor _
              Asc(Mid$(Code, intLocation%, 1)))
  Next i%
End Sub

Public Sub loginfailcheck(loginfailID As String)
Dim usRblock As ADODB.Recordset
If db Is Nothing Then dbconnect

failname = loginfailID
If failname = failname Then
   If loginfail = 3 Then
     Set usRblock = New ADODB.Recordset
       usRblock.Open ("Update T_users_acct set acctStatus = 'Locked' where userID='" & failname & "'"), db, 3, 3
       Call UserLogs(acctid, "Blocked")
       MsgBox "Account has been locked please contact your system administrator     ", vbInformation
       Set usRblck = Nothing
   End If
Else
   loginfail = 0
End If
Set db = Nothing

End Sub

Public Sub Lock_UnlockAcct(ByRef idNo As Long, UserP As String, SelID As Double, act As String)
Dim updte As ADODB.Recordset
Set updte = New ADODB.Recordset

dbconnect
With updte
     .Open ("select IDnumber, userID, userPassword from T_users_acct where IDnumber = " & CDbl(idNo) & " "), db, 3, 3
     EncryptDecrypt (UserP)
     UserP = temp
     If UserP = .Fields("userPassword") Then
        Set updte = Nothing
        .Close
        
        Set updte = New ADODB.Recordset
        .Open ("Update T_users_acct set acctStatus = '" & act & "' where IDnumber=" & SelID & " "), db, 3, 3
        MsgBox "Account Update Successful!   ", vbInformation
        
     Else
        MsgBox "Invalid Password!   ", vbCritical
        
     End If
    
End With
Set updte = Nothing
Set db = Nothing

End Sub

Public Sub UserLogs(idNum As Long, activity As String)
Dim UserLog As ADODB.Recordset
Set UserLog = New ADODB.Recordset
If db Is Nothing Then dbconnect

UserLog.Open ("insert into T_user_logs values(" & CDbl(idNum) & ", '" & activity & "', '" & Time & "', '" & Date & "') "), db, 3, 3
Set UserLog = Nothing
Set db = Nothing

End Sub

Public Sub CompactRepairAccessDB()
Dim objJRO As JRO.JetEngine
Dim strConnSource As String
Dim strConnDestination As String

dbpath = DataBasedbpath & "Database\Data.mdb"
Compactedfile = DataBasedbpath & "Database\DataCpd.mdb"

If db Is Nothing Then
   
Else
   db.Close
End If

mypassword = "TechServeEra"
'set connection string info
strConnSource = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath & ";" & _
                "Jet OLEDB:Database Password=" & mypassword
strConnDestination = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Compactedfile & ";" & _
                     "Jet OLEDB:Database Password=" & mypassword
    'initiate object
Set objJRO = New JRO.JetEngine
    'compact database from source to destination
objJRO.CompactDatabase strConnSource, strConnDestination
    'release objects
Set objJRO = Nothing
    'Delete the old file
Kill dbpath
Name Compactedfile As dbpath
dbconnect
MsgBox "Data Base has been compacted"

End Sub

