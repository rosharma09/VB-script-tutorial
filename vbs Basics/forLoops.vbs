' Program to illustrate FOR loops in VBS 

Option Explicit 

On Error Resume Next 

Const SITE_NAME = "www.testdummy.com"

Dim a 

' For loop syntax : For i = value To Value 

For a= 1 To 10 
	MsgBox "The value of a is : " &a , 0 , SITE_NAME 
	Next
	

Dim b 
For b = 2 To 20 Step 2
	MsgBox "2*" &b& "=" &(2*b) , 0 , SITE_NAME
	Next
	
Dim Counter , i , j 

i = 10 
j = 5

For counter = i To j Step -1
	MsgBox "Hello" &i , 0 , SITE_NAME
	Next
	
