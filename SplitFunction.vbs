'============================================
' Program to demonstrate the split function 
'============================================

Option Explicit 
On Error Resume Next 

' Declare a variable 

Const SITE_NAME = "www.google.com"

' Declare an array to store the split array 

Dim splitArray
splitArray = Split(SITE_NAME , ".")

MsgBox splitArray(0)
MsgBox splitArray(1)
MsgBox splitArray(2)