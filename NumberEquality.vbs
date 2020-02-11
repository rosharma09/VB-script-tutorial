' Program to find the entered numbers are equal or not 

Option Explicit 
On Error Resume Next 

Dim a , b , num

Const SITE_NAME = "www.euality.com"

' User input values 
a = InputBox("Enter 1st Number : ")
b = InputBox("Enter 2nd Number : ")

' Logic to find the equality 

If CInt(a) = CInt(b) Then 
	MsgBox "The entered numbers are equal" , 0 , SITE_NAME
End If

' Multiple conditions

num = 20

If num < 20 Then
	MsgBox "Number is lesser than 20" , 0 , SITE_NAME
End If
	
If num = 20 Then 
	MsgBox "Number is equal to 20" , 0 , SITE_NAME
End If

If num > 20 Then 
	MsgBox "Number is Greater than 20" , 0 , SITE_NAME
End If
