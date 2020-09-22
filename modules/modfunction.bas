Attribute VB_Name = "Modfunction"




Public Function Louie(strng As String, loc As CursorLocationEnum, connectectionString As String) As Recordset
  Set Louie = New Recordset
  Louie.CursorLocation = loc
 Louie.CursorType = adOpenKeyset
  Louie.LockType = adLockOptimistic
Louie.Open strng, connectectionString, , , adCmdText
End Function



