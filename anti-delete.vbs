Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("C:\testfile.txt")

'Set the system attributes of the file
objFile.Attributes = 2 + 4 + 32

'Disallow access to the file for all users
objFile.Permissions = 0

'Lock the file
objFile.Lock

'Read the file and store the result in a variable
strText = objFile.ReadAll

'Unlock the file
objFile.Unlock

'Close file
objFile.Close

'Delete object file
Set objFile = Nothing

'Выводим сообщение о том, что файл успешно защищен
WScript.Echo "The file has been successfully protected from deletion."

'In this script, we create a text file on the C:\ drive and set its system attributes, which prevent deletion, renaming, or modification of the file. 
'We also disallow access to this file for all users and lock it to prevent other processes from changing its content or properties.

'It is important to note that this script can only prevent file removal on a computer with public access. 
'If the computer already has system administration capabilities, this script can be executed with administrative rights, 
'allowing the file to be deleted.