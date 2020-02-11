'======================================
' Different ways of declaring an array
'======================================

Option Explicit
On Error Resume Next

' Declare variable 

Const SITE_TEXT = "www.dummysite.com"

' Declare an array named A with 10 elements 
Dim ArrayTest ' Declare a variable
ArrayTest = Array("user1" , "user2" , "user3" , "user4") ' intializing the array

Redim Preserve ArrayTest(9)

ArrayTest(8) = "user8"
ArrayTest(9) = "user9"

MsgBox ArrayTest(0) , 0 , SITE_TEXT
MsgBox ArrayTest(8) , 0 , SITE_TEXT
MsgBox ArrayTest(9) , 0 , SITE_TEXT