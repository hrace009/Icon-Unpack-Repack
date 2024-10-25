Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set AdodbStream = CreateObject("Adodb.Stream")
Set oImg = CreateObject("WIA.ImageFile")
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
exe=scriptDir&"montage.exe"
input=WScript.arguments(0)
icontmp=scriptDir&objFSO.GetBaseName(input)&"-tmp"
icondds=scriptDir&objFSO.GetBaseName(input)&".dds"
icontxt=scriptDir&objFSO.GetBaseName(input)&".txt"

If objFSO.FolderExists(icontmp) Then objFSO.DeleteFolder(icontmp)
objFSO.CreateFolder(icontmp)

For Each i in objFSO.GetFolder(input).files
  a=a+1
  If a=1 Then
    oImg.LoadFile i.path
    w=oImg.Width
    h=oImg.Height
  End If
Next

m=128

k=a\m
r=a mod m
If r>0 Then k=k+1

set g = objFSO.OpenTextFile(icontxt, 2, True, True)
g.WriteLine w
g.WriteLine h
g.WriteLine k
g.WriteLine m
For Each i in objFSO.GetFolder(input).files
  bn=bn+1
  rn=100000+bn
  b=objFSO.GetBaseName(i)
  e=objFSO.GetExtensionName(i)
  If b="-unknown.dds" Then b="unknown.dds"
  g.WriteLine b
  tmpfile=icontmp&"\"&rn&"."&e
  objFSO.GetFile(i).copy(tmpfile)
Next
g.close

AdodbStream.Charset = "unicode"
AdodbStream.Open
AdodbStream.LoadFromFile icontxt
g2u = AdodbStream.ReadText
AdodbStream.Close
AdodbStream.Charset = "gb2312"
AdodbStream.Open
AdodbStream.WriteText g2u
AdodbStream.SaveToFile icontxt,2
AdodbStream.Close

input=Chr(34)&icontmp&Chr(34)
icondds=Chr(34)&icondds&Chr(34)
exe=Chr(34)&exe&Chr(34)

WshShell.run ""&exe&" "&input&"\*.* -geometry +0+0 -tile "&m&"x"&k&" -background none -define dds:compression=none -define dds:cluster-fit=false "&icondds&"",0,True

objFSO.DeleteFolder(icontmp)
WScript.Echo "done"