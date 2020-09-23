Attribute VB_Name = "globalfunctions"
Function Random(Lowerbound As Long, Upperbound As Long) As Long
Randomize
Random = Int((Upperbound - Lowerbound) * Rnd + Lowerbound)

End Function

Function checkPermission(ByRef ConCol As String, ByRef Index As Integer, idNum As Double) As Boolean
Dim perm As ADODB.Recordset
Set perm = New ADODB.Recordset
If db Is Nothing Then dbconnect

With perm
     .Open ("select * from T_acct_restric where IDnumber= " & CDbl(idNum) & " "), db, 3, 3
       If ConCol = "tab" Then
          If CBool(.Fields(6)) = True Then
             checkPermission = True
          Else
             checkPermission = False
          End If
       End If
       
       If ConCol = "combo" Then
          If CBool(.Fields(5)) = True Then
             checkPermission = True
          Else
             checkPermission = False
          End If
       End If
       
       If ConCol = "btn" Then
          If CBool(.Fields(Index)) = True Then
             checkPermission = True
          Else
             checkPermission = False
          End If
       End If
     .Close
     Set perm = Nothing
End With

End Function


