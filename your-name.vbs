Dim answer, answer2, arrUserData

arrUserData = Array(0,1)

 

answer = MsgBox("Ready to start?", vbYesNo, "Welcome")

 

If answer = vbYes Then

'Asks the User for input their name and age

Do

arrUserData(0) = InputBox("What's your name?" & VbCrLf & VbCrLf & VbCrLf & "(If you want, see what happens when you leave this empty)")

If arrUserData(0) = "" Then

MsgBox("Please, enter your name")

End If

Loop While arrUserData(0) = ""

Do

arrUserData(1) = InputBox("What's your age?")

If arrUserData(1) = "" Then

MsgBox("Please, enter your age")

End If

Loop While arrUserData(1) = ""

 

'Prints the name and age the User just wrote

answer2 = MsgBox("Welcome " & arrUserData(0) & " of " & arrUserData(1) & " years old! Very nice meeting you!" & VbCrLf & VbCrLf & VbCrLf & "Do you wish to continue?", vbYesNo)

If answer2 = vbYes Then

MsgBox("We'll continue VBscripting tomorrow, bye now")

Else MsgBox("That's ok, bye bye!")

End If

 

Else MsgBox("That's ok, bye bye!")

End If
