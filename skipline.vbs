Const ForReading = 1
Const ForWriting = 2
Set objFSO = CreateObject("Scripting.FileSystemObject")
ReadFileName  = wscript.arguments(0)
LinesToSkip   = wscript.arguments(1)
WriteFileName = wscript.arguments(2)

' for demo purposes; can delete
'--------------------------------------
StartRead = Timer
'--------------------------------------

Set ReadFile  = objFSO.OpenTextFile (wscript.arguments(0),  ForReading)
strText = ReadFile.ReadAll
Readfile.Close

' for demo purposes; can delete
'--------------------------------------
EndRead = Timer
ReadTime = EndRead - StartRead
StartSplit = Timer
'--------------------------------------

ArrayOfLines = Split(strText, vbLf)

' for demo purposes; can delete
'--------------------------------------
Endsplit = Timer
SplitTime = Endsplit - StartSplit
StartWrite = Timer
'--------------------------------------

Set Writefile = objFSO.CreateTextFile (WriteFileName, ForWriting)
For j = LinesToSkip To UBound(ArrayOfLines)
	Writefile.writeline ArrayOfLines(j)
Next
Writefile.Close

' for demo purposes; can delete
'--------------------------------------
EndWrite = Timer
TotalTime = EndWrite - StartRead
WriteTime = EndWrite - StartWrite
wscript.echo "Number of Lines in file: " & UBound(ArrayOfLines)
wscript.echo "Read file in             " & ReadTime  & " sec(s)"
wscript.echo "Split file in            " & SplitTime & " sec(s)"
wscript.echo "Wrote output file in     " & WriteTime & " sec(s)"
wscript.echo "Total time taken         " & TotalTime & " sec(s)"
'--------------------------------------
