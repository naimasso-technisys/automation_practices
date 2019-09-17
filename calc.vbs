dim arrFruits, arrMath1, arrMath2, Choice, num1, num2, result, str1, str2

 

arrFruits = Array("Apples","Oranges","Bananas")

arrMath1 = Array(1,3,5)

arrMath2 = Array(1,2,3)

num1 = UBound(arrMath1)

num2 = UBound(arrMath2)

 

For x=0 To num1

result = arrMath1(x) + arrMath2(x)

MsgBox(arrMath1(x) & " + " & arrMath2(x) & " = " & result)

Next

 

str1 = Join(arrMath2,", ")

MsgBox("Your array contains these values: " & str1)

 

str2 = Join(arrFruits,", ")

MsgBox("You need to eat more: " & str2)

 

Choice = MsgBox("Do you want to exit?", vbYesNo, "Continue")

If Choice = vbYes Then

MsgBox("Have a nice day! :)")

Else MsgBox("Great! Let's continue VBscripting :)")
