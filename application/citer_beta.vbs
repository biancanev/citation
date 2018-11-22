dim URL
URL=InputBox("Please insert the url...", "URL Wizard")
set xmlhttp = createobject ("msxml2.xmlhttp.3.0")
xmlhttp.open "get", URL, false
xmlhttp.send
strOutput = split(xmlhttp.responseText,"title")(1)
strT = split(strOutput, "</title>")(0)
StrTi = Replace(strT, "<", "")
StrTitleBeta = Replace(strTi, "/", "")
strTitleAlpha = Replace(strTitleBeta, "content", "")
strTitleDelta = Replace(strTitleAlpha, "=", "")
strTitleGamma = Replace(strTitleDelta, chr(34), "")
strTitleZebra = Replace(strTitleGamma, "meta property:", "")
StrTitle = Replace(strTitleZebra, ">", "")
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
curDate =  Month(Date) & "/" & Day(Date) & "/" & Year(Date)
msgBox "Title: " & strTitle & Chr(13) & Chr(10) & "Author: "&strAuthor1 & " " & strAuthor2 & Chr(13) & Chr(10) & "Date Accessed: " & curDate & Chr(13) & Chr(10) & "Date Published: " & Chr(13) & Chr(10) & "URL: " & URL , VBOKOnly, "Here are your results:" 