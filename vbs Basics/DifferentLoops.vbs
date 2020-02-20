' This program is to illustrate different loops
' Loops : Do While , Do Until
' wscript.quit
' exit do

' Define the variables

' Example for Do loop

Option Explicit 
On Error Resume Next 

Dim a 
a = 0 
Do
	a = a + 1
	MsgBox  a& " : Im in a loop"
	If a = 10 Then
		exit do ' we can use exit do to terminate the loop
		End If
Loop

MsgBox "End of the loop"

'Example for Do Until

Dim b
b = 0

Do Until b = 5 ' In case of Do Until, the loop executes till the condition is satisfied  - Exit on true 
	MsgBox "b =" &b
	b = b + 1 
Loop

Dim Test , UserName , PassWord

Dim LoginCount

Test = "Fail"
LoginCount = 0

Do Until Test = "Pass" 

	UserName = InputBox("Enter the username : ")
	PassWord = InputBox("Enter the password : ")
	
	If (UserName = "rosharma" AND PassWord = "Reg!234") Then
		Test = "Pass"
	Elseif (UserName <> "rosharma" OR PassWord <> "Reg!234") Then
			LoginCount = LoginCount + 1
			If (LoginCount = 2) Then 
			MsgBox "User reached maximum retries" , 0
			MsgBox "User account is locked" , 0 
			Wscript.quit
			End If
	End If
Loop

' Example for Do While loop 

Dim x 
x = 0 

Do While x < 5 ' Exit when false
	MsgBox "Value of x is : " &x , 0 
	x = x + 1
	MsgBox "Incremented value of x :" &x , 0
Loop
	



