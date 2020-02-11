' Program to find the entered numbers are equal or not 

Option Explicit 
On Error Resume Next 

Dim a , b 

Const SITE_NAME = "www.euality.com"

' User input values 
a = InputBox("Enter 1st Number : ")
b = InputBox("Enter 2nd Number : ")

' Logic to find the equality 

If CInt(a) = CInt(b) Then
	MsgBox "The entered numbers are equal" , 0 , SITE_NAME
End If
	
	

