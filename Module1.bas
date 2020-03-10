Attribute VB_Name = "Module1"

Sub main()

Dim c  As Object
Set c = Range("b1")
For Each c In Selection
  Call getInfo(c)
Next

End Sub


Sub getInfo(c As Range)

Dim url
Dim buf, removeEnterTab
url = c.Value

Dim HTTP
Set HTTP = CreateObject("msxml2.xmlhttp")
HTTP.Open "GET", url, False
HTTP.send
buf = HTTP.responsetext
removeEnterTab = Replace(buf, Chr(13), "")
removeEnterTab = Replace(buf, Chr(10), "")
'Range("c1") = removeEnterTab

Dim targetCollection
Set targetCollection = getString(removeEnterTab, " <p class=""join_fee"">", "</p>")
targetStr = Replace(targetCollection(0), " ", "")
Debug.Print targetStr
c.Offset(0, 1).Value = targetStr

Set targetCollection = getString(removeEnterTab, "<p class=""ymd"">", "</p>")
targetStr = Replace(targetCollection(0), " ", "")
Debug.Print targetStr
c.Offset(0, 2).Value = targetStr

Set targetCollection = getString(removeEnterTab, "<span class=""hi"">", "</span>")
targetStr = Replace(targetCollection(0), " ", "")
Debug.Print targetStr
c.Offset(0, 3).Value = targetStr

Set targetCollection = getString(removeEnterTab, "<span class=""amount"">", "</span>êl")
targetStr = Replace(targetCollection(0), " ", "")
Debug.Print targetStr
c.Offset(0, 4).Value = targetStr

Set targetCollection = getString(removeEnterTab, "<p class=""ptype_name"">", "</p>")
targetStr = Replace(targetCollection(0), " ", "")
Debug.Print targetStr
c.Offset(0, 5).Value = targetStr

End Sub


Function getString(str, exp1, exp2, Optional globalBool = False)

Dim RE, Matches As Variant
Set RE = CreateObject("vbscript.regexp")
With RE
  .Pattern = exp1 & ".*?" & exp2
  .Global = globalBool
  .ignorecase = True
End With
Set Matches = RE.Execute(str)
Set getString = Matches

End Function

