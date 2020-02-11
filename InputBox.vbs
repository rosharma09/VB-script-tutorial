'=======================================================
' Program to illustrate InputBox to take input from user
'=======================================================

'  Declare the properties

Option Explicit
On Error Resume Next 

Const SITE_TITLE = "www.simpleAddition.com"

Dim input1 , input2 , name , sum

' Getting input from user 

name = InputBox("Please enter your name :")

input1 = InputBox("Enter the first Number : ")

input2 = InputBox("Enter the second number : ")

sum = input1 + input2 ' the issue here is that the user entered value is taken as a string and + operation concatenates the two strings

' Display the message 

MsgBox " Hello " &name& " welcome to the simple calculator " , 0 , SITE_TITLE
MsgBox "The sum of " &input1& " and " &input2& " is : " &sum , 0 , SITE_TITLE  


' To resolve this we are going to use CInt to convert the string to integer value 

' CInt is a function that converts string into integer in VBS 

Dim x , y , z 

x = InputBox("Enter x : ")
y = InputBox("Enter y : ")

z = CInt(x) + CInt(y)

MsgBox "x+y = " &z , 0 , SITE_TITLE

' Customized InputBox example
' we can add some customization to our InputBox to display the title and default text and alignment

Const LOGIN_PAGE = "www.login.com"

const HOME_PAGE = "wwwm.homepage.com"

Dim userName , Password 
' InputBox("Text in the text box" , "Title of the page" , "predefined text in the input box" , X-cordinate , Y-cordinate)
userName = InputBox("Enter the username" , LOGIN_PAGE , "username" , 5000 , 5000)
Password = InputBox("Enter the password" , LOGIN_PAGE , "password" , 5000 , 5000)

MsgBox "Welcome to our world " , 0 , HOME_PAGE



