Dim win
win = InputBox("WIN? (type no to reset)", "Win Counter v0.9(BETA) slash awesome", "YES")
Set objFS = CreateObject("Scripting.FileSystemObject")
strFile = "wins.txt"
Set objFile = objFS.OpenTextFile(strFile)
Dim count
count = objFile.ReadLine
Dim newCount
newCount = count + 1
objFile.Close
Set objFile = objFS.GetFile(strFile)
objFile.Delete
Set objFile = objFS.CreateTextFile(strFile, True)
If win = "no" Then newCount = 0
WScript.Echo newCount & " Wins today!"
objFile.Write(newCount)
objFile.close