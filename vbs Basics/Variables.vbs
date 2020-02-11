'===============================================================
' This script is to demonstrate :
' variables , constants , MsgBox , Commands , line continuation
' option explict , On Error Resume Next
'===============================================================

' Option Explict is used to enforce the declaration of a variable in a script
' Its a good practie to use Option Explicit to avoid confusion in case of large script
'Option Explicit

' On Error Resume Next is used to skip the line containing error and proceed to the next line
On Error Resume Next

' Declare variables 

' Its a good practice to declare same type variables in the same line
Dim num1 , num2 
Dim sum
Dim name , city

' We can declare a constant value in the script using Const keyword and use caps for constant variables in the script
Const SITE_NAME = "www.testsite.com"

' Intialization of the variables

num1 = 100
num2 = 200

' Processing of the variables

sum = num1 + num4

' Display the message boxes 

MsgBox sum

MsgBox " Sum of " &num1& " and " &num2& " is equal to " &sum1 , 0 , SITE_NAME

' we can also use line continuation using _

MsgBox "++Line Continuation example : Thee sum of " &num1&_
				"and " &num2& "is " &sum , 0 , SITE_NAME	


' In case if we use a variable that is not defined in the script we are going to get runtime error
name1 = "Rohan"
city = "Bangalore"

MsgBox name1 & " Lives in " & city & " !!!! " , 0 , SITE_NAME


