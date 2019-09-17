Dim question1, question2, question3, str

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set WshShell = WScript.CreateObject("WScript.Shell")

 

objFSO.CreateTextFile("C:\Users\Public\Scripts_test.txt")

question1 = MsgBox("Please check that 'Scripts_test.txt' was succesfully created on your system. File location: C:\Users\Public\" & VbCrLf & VbCrLf & "Is the file emtpy?", vbYesNo)

                If question1 = vbYes Then

                Set txtFile = objFSO.OpenTextFile("C:\Users\Public\Scripts_test.txt", 8, True)

                txtFile.WriteLine("Welcome! This is a .txt file created for VBscript practice purposes. I hope you have a nice learning experience with VBscript!")

                txtFile.Close

                               question2 = MsgBox("Now, please close the file and re-open it." & VbCrLf & VbCrLf & "Is the file empty now?", vbYesNo)

                                               If question2 = vbNo Then

                                               MsgBox("Wow! Who wrote that? ;)")

                                               Else If question2 = vbYes Then

                                                               Do

                                                               question3 = MsgBox("Are you sure? Check again please!" & VbCrLf & VbCrLf & "Is it still empty?", vbYesNo)

                                                               Loop While question3 = vbYes

                                                               If question3 = vbNo Then

                                                                               MsgBox("Nice, hu? ;)")

                                                               End If

                                               End If

                               End If

                End If

 

MsgBox("The file will be deleted now. Have a nice day")

If objFSO.FileExists("C:\Users\Public\Scripts_test.txt") Then

   objFSO.DeleteFile("C:\Users\Public\Scripts_test.txt")

End If
