' Program to find the greatest of two numbers

Option Explicit
On Error Resume Next

' Declare the variables 

Dim Num1 , Num2 

Const SITE_NAME = "test site"

Num1 = InputBox("Enter Num1 : ")
Num2 = Inputbox("Enter Num2 : ")


If CInt(num1) > CInt(Num2) Then 
	MsgBox Num1& " > " &Num2 , 0 , SITE_NAME
ElseIf CInt(num2) > CInt(Num1) Then 
	MsgBox Num2& " > " &Num1 , 0 , SITE_NAME
End If

' Another example 

Dim count 
count = 1

If count = 0 Then
	MsgBox "Count value is 0" , 0 , SITE_NAME
ElseIf count = 1 Then 
	MsgBox "Count value is 1" , 0 , SITE_NAME
ElseIf count = 2 Then
	MsgBox "Count value is 2" , 0 , SITE_NAME
End If

' To find the greatest of three numbers

Dim x , y , z

x = InputBox("Enter the 1st Number :")
y = InputBox("Enter the 2nd Number :")
z = InputBox("Enter the 3rd Number :")

' We can make use of the operators AND to have multiple conditions in IF condition

If CInt(x) > CInt(y) AND CInt(x) > CInt(z) Then	
	MsgBox "The largest of " &x& " , " &y& " and " &z& " is :" &x , 0 , SITE_NAME
ElseIf CInt(y) > CInt(x) AND CInt(y) > CInt(z) Then	
	MsgBox "The largest of " &x& " , " &y& " and " &z& " is :" &y , 0 , SITE_NAME
ElseIf CInt(z) > CInt(y) AND CInt(z) > CInt(x) Then	
	MsgBox "The largest of " &x& " , " &y& " and " &z& " is :" &z , 0 , SITE_NAME
End If


' String match 

Dim name1 , name2 

name1 = "john"
name2 = "peter"

If name1 <> name2 Then
	MsgBox "Names did not match"
End If

' Multiple conditions 

Dim x1 , x2 , x3 

x1,x2,x3 = 1

If x1 = x2 AND x2 = x3 Then 
	MsgBox "Both the conditions are satisfied"
End If

' use of OR in IF condition 

If x1 = x2 OR x2 <> x3 Then 
	MsgBox "One of the condition is true"
End If 
