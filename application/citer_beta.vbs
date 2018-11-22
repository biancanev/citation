dim URL
URL=InputBox("Please insert the url...", "URL Wizard")
set xmlhttp = createobject ("msxml2.xmlhttp.3.0")
xmlhttp.open "get", URL, false
xmlhttp.send
strOutput = split(xmlhttp.responseText,"<title>")(1)
strTitle = split(strOutput, "</title>")(0)
strAuth = split(xmlhttp.responseText,"by ")(1)
strAuthor1 = split(strAuth, " ")(0)
strAuthor2 = split(strAuth, " ")(1)
If strAuthor1 = "" then
strAuth = split(xmlhttp.responseText, "<img")(1)
strAuthor=split(strAuth, ">")(1)
strAuthor1=split(strAuthor, " ")(0)
strAuthor2=split(strAuthor, " ")(1)
End If
If strAuthor1 = "" Then
strAuth = split(xml.http.responseText, "<hr>")
strAuthor = split(strAuth, "<")(1)
strAuthor1 = split(strAuthor, " ")
End If
domain=split(URL, "https://www.")(1)
domainname=split(domain, ".")(0)
If InStr(domainname, "cnn") Then
publisher="CNN"
ElseIf InStr(domainname, "google") Then
publisher="Google"
ElseIf InStr(domainname, "yahoo") Then
publisher="Yahoo"
Else
publisher=domainname
End If
curDate =  Month(Date) & "/" & Day(Date) & "/" & Year(Date)
mlaDate = Day(Date) & " " & Month(Date) & " " & Year(Date)
msgBox "Title: " & strTitle & Chr(13) & Chr(10) & "Author: "&strAuthor1 & " " & strAuthor2 & Chr(13) & Chr(10) & "Date Accessed: " & curDate & Chr(13) & Chr(10) & "Date Published: " & Chr(13) & Chr(10) & "Publisher: " & publisher & Chr(13) & Chr(10) & "URL: " & URL , VBOKOnly, "Here are your results:" 
msgBox strAuthor2 & ", " & strAuthor1 & ". " & Chr(34) & strTitle & "." & Chr(34) & publisher & ", "& publisher & " Inc., " & mlaDate & " " & URL, VBOkOnly, "MLA format:"