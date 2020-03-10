Attribute VB_Name = "Module2"
Sub getURLs()

Dim SH As Worksheet
Set SH = Worksheets("sheet2")

SH.Cells.ClearContents

Dim url
Dim buf, removeEnterTab
url = "https://connpass.com/calendar/"

Dim HTTP
Set HTTP = CreateObject("msxml2.xmlhttp")
HTTP.Open "GET", url, False
HTTP.send
buf = HTTP.responsetext
removeEnterTab = Replace(buf, Chr(13), "")
removeEnterTab = Replace(buf, Chr(10), "")
'Range("c1") = removeEnterTab
Dim targetStr As Variant
Set targetStr = getString(removeEnterTab, "<div class=""main_area mt_20"">", "</table>")

Dim urls
Set urls = getString(targetStr(0), "href=""", """>", True)

For Each c In urls
  Dim temp
  temp = c
  temp = Replace(temp, "href=""", "")
  temp = Replace(temp, """ target=""_blank"">", "")
  SH.Range("a10000").End(xlUp).Offset(1, 0) = temp
Next c

End Sub

