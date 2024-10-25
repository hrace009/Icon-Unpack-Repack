Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set AdodbStream = CreateObject("Adodb.Stream")
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
exe=scriptDir&"magick.exe"
icontxt=WScript.arguments(0)
icondds=scriptDir&objFSO.GetBaseName(icontxt)&".dds"
temptxt=scriptDir&objFSO.GetBaseName(icontxt)&"-temp.txt"
output=scriptDir&objFSO.GetBaseName(icontxt)
If not objFSO.FolderExists(output) Then objFSO.createFolder(output)

AdodbStream.Charset = "gb2312"
AdodbStream.Open
AdodbStream.LoadFromFile icontxt
g2u = AdodbStream.ReadText
AdodbStream.Close
AdodbStream.Charset = "unicode"
AdodbStream.Open
AdodbStream.WriteText g2u
AdodbStream.SaveToFile temptxt,2
AdodbStream.Close

icondds=Chr(34)&icondds&Chr(34)
exe=Chr(34)&exe&Chr(34)

set f = objFSO.OpenTextFile(temptxt,1,False,True)
Do until f.AtEndOfStream
  a=f.ReadLine
  b=Trim(a)
  If b<>Empty Then
    bn=bn+1
    dn=bn-4
    If bn=1 Then w=b
    If bn=2 Then h=b
    If bn=4 Then m=b
    If bn>4 Then
      x=dn mod m
      y=dn\m
      If x=0 Then
        x=m-1
        y=y-1
       Else
        x=x-1
      End If
      x=x*w
      y=y*h
      If b="unknown.dds" Then b="-unknown.dds"
      outputimage=output&"\"&b&".png"
      outputimage=Chr(34)&outputimage&Chr(34)
	  WshShell.run ""&exe&" "&icondds&" -strip -alpha on -crop "&w&"x"&h&"+"&x&"+"&y&" "&outputimage&"",0,False
      WScript.Sleep 200
    End If
  End If
Loop
f.close

objFSO.DeleteFile(temptxt)
WScript.Echo "done"