

'# $language = "VBScript"

'# $interface = "1.0"

 

'Written by: Alexander M. Meyers

'Last Edited: 3/21/2019 12:37AMs

'Version 1.0

Dim wShell

Dim oExec

Dim sFileSelected

Dim filename

Dim filenameposition

Dim confirmresult

Dim localip

Dim strFile

Dim defaultpath

Dim objFile

Dim detectedpath

Dim sourcepath

Dim filemoved

Dim transferfailed

Dim loopbreaker



Sub Main

filemoved = 0

'We first check the settings on the SecureCRT Client we're using to ensure

'that we actually have stuff set correctly to use this script.

'Otherwise we throw an error.

 

'Proceed to use the path parsed from the Global.ini file and store it inside a temp variable.


'First need to grab the username so we know where to look

Set wshNetwork = CreateObject("wscript.Network")

strUserName = wshNetwork.UserName

strFile = "C:\Users\" & strUserName & "\AppData\Roaming\VanDyke\Config\Global.ini"

defaultpath = "C:\Users\" & strUserName & "\Desktop\tftp"

 

 

'we use the Microsoft HTML Applications executable available on all MS installs to run an

'activeX object which allows us to request a file browser pop up to search the file system

'for our file to transfer.

Do

Set wShell=CreateObject("WScript.Shell")

Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE> <title> Select a file to transfer </title> <script>FILE.click(); new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

sFileSelected = oExec.StdOut.ReadLine

 

'from the provided path, we parse out our file name and then pass this to the xmodem string to copy

filenameposition = InStrRev ( sFileSelected, "\", -1, vbTextCompare)

filenameposition = Len(sFileSelected) - filenameposition

filename = Right ( sFileSelected, filenameposition )

filenameposition = InStrRev ( sFileSelected, "\", -1, vbTextCompare)

sourcepath = Left(sFileSelected, Len(sFileSelected) - Len(filename))

 

 

If StrComp(filename,"") = 0 Then

'do nothing

Else

confirmresult = MsgBox ("Are you sure this is the file you want to transfer? " & vbCr & "File Name: " & filename & vbCr & "File path: " & sFileSelected, vbYesNoCancel, " Confirm")

End If

Loop Until (confirmresult <> vbNo or StrComp(filename,"") = 0)

 

'this is the part where having a box to test with will come in handy

'but for the time being we're just using the sample code provided by

'the devs of SecureCRT

'This is commented out for testing purposes at the moment.

 

If confirmresult = vbYes Then

localip = InputBox ("Please input your TFTP target. If not connected to the internet, this would be the IP to your laptop. DON'T USE THE CPE")

 

If StrComp(localip,"") = 0 Then

Exit Sub 'User cancelled.

End If

 

'Now we pull the download/upload directories from the Global.ini file

'We throw an error if these are not already configured.

 

Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(strFile,1)

        'Read the entire file into memory.

    strFileText = objFile.ReadAll

        'Close the file.

    objFile.Close

        'Split the file at the new line character. *Use the Line Feed character (Char(10))

    arrFileText = Split(strFileText,Chr(10))

    For i = LBound(arrFileText) to UBound(arrFileText)

            'If the line is not blank process it.

        If StrComp(arrFileText(i), "") <> 0 Then

                                contents = arrFileText(i)

                                contents = Replace(contents, vbCr, "")

                                contents = Replace(contents, vbLf, "")

                                arrFileText(i) = contents

                                If StrComp(Left(arrFileText(i), Len("S:""TFTP Server Download Directory""=")),"S:""TFTP Server Download Directory""=") = 0 Then

                                                                'Check Global Settings to see if the TFTP server is enabled and pull the path.

                                                                'MsgBox "Line Found"

                                                                If StrComp(Right(arrFileText(i),(Len(arrFileText(i)) - Len ("S:""TFTP Server Download Directory""="))), "") <> 0 Then

                                                                                'MsgBox "Folder detected " & Right(arrFileText(i),(Len(arrFileText(i)) - Len ("S:""TFTP Server Download Directory""=")))

                                                                                detectedpath = Right(arrFileText(i),(Len(arrFileText(i)) - Len ("S:""TFTP Server Download Directory""="))) & "\"

                                                                                changeshappened = 1

                                                                               

                                                                End If

                                End If

                End If

 

                Next

objFile = Null

'end of reading from Global.ini

 

'Once we have determined the location of the download folder, we compare its location to our file location

'and move to the file to the correct folder before beginning the transfer if it is not already in the right folder.

 

If StrComp(detectedpath, sourcepath) <> 0 Then
MsgBox "file not in correct location, moving temporarily to SecureCRT download folder"

Set objFile = CreateObject("Scripting.FileSystemObject")

If objFile.FileExists(sFileSelected) = True Then

objFile.MoveFile sFileSelected, detectedpath

filemoved = 1

Else

MsgBox "File doesnt exist anymore! Terminating."

Exit Sub

End If

End If

 
crt.Screen.Send vbCR
crt.Screen.Send vbCR
crt.Screen.Send vbCR

crt.Screen.Send "copy tftp: nvram:" & vbCR

If crt.Screen.WaitForString("Address", 3, True) <> True Then
MsgBox "Transfer Failed"
transferfailed = 1
End If

crt.Screen.Send localip & vbCR

crt.Screen.WaitForString "Source", 3, True

crt.Screen.Send filename & vbCR

crt.Screen.WaitForString "Destination", 3, True

crt.Screen.Send vbCR

If crt.Screen.WaitForString( "over write", 15, True) = True Then
crt.Screen.Send "y" & vbCR
End If

If crt.Screen.WaitForString( "Loading", 30, True) <> True or transferfailed = 1 Then
MsgBox "Transfer failed"
If filemoved = 1 Then

Set objFile = CreateObject("Scripting.FileSystemObject")

objFile.MoveFile detectedpath & filename, sourcepath

Else
Exit Sub
End If
End If

loopbreaker = 0


Do
crt.screen.Synchronous = True
loopbreaker = crt.Screen.ReadString( "copied", "secs", "out", "error", "%Error", "timed", "interrupted", 3, True)
'Msgbox loopbreaker
If crt.Screen.MatchIndex = 1 or crt.Screen.MatchIndex = 2 Then
MsgBox "Transfer successful"
ElseIf crt.Screen.MatchIndex <> 0 Then
MsgBox "Transfer failed"
End If
Loop Until crt.Screen.MatchIndex <> 0
 

Else

'Do nothing

End If

If filemoved = 1 Then

Set objFile = CreateObject("Scripting.FileSystemObject")

objFile.MoveFile detectedpath & filename, sourcepath

End If

End Sub