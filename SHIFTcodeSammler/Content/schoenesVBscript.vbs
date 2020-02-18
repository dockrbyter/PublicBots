Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\m6xim\Documents\UiPath\SHIFTcodeSammler\Content\file.txt",1)
strFileText = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

strFileText = Replace(strFileText, vbCr, " ")
strFileText = Replace(strFileText, vbLf, " ")



Dim objRegExp
Set objRegExp = New RegExp

objRegExp.IgnoreCase = False
objRegExp.Global = True
objRegExp.Pattern = "VIP (EMAIL|VAULT).+?POINTS  (.+?)  Redeem"


Dim objMatch

Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\m6xim\Documents\UiPath\SHIFTcodeSammler\Content\fileout.csv",2,true)
For Each objMatch in objRegExp.Execute(strFileText)
   objFileToWrite.WriteLine(objMatch.SubMatches(1) & ";" & objMatch.SubMatches(0))
Next
objFileToWrite.Close
Set objFileToWrite = Nothing
