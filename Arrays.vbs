'===========================================================
' Arrays in VB scripts 
' we can implement 1D and 2D arrays in our VB script
'===========================================================

'=================
' 1D array
'=================

Option Explicit

'On Error Resume Next

' Declare variable

Dim num1 

' Declare a 1D array named array1D with 5 elements 

Dim array1D(4)

Const SITE_TITLE = "www.TestArray.com"

' Assigning values to the elements in array

array1D(0) = 100
array1D(1) = 200
array1D(2) = 300
array1D(3) = 400
array1D(4) = 500

' Assign the value of an array element to a variable

num1 = array1D(1)

MsgBox "Value of num1 is :" &num1 , 0 , SITE_TITLE

MsgBox "Value of array1D(3) is :" &array1D(3) , 0 , SITE_TITLE

MsgBox "Value of array1D(4) is :" &array1D(4) , 0 , SITE_TITLE


'=================
' 2D array
'=================

' Declare a 2D array named array2D to store name and favorite fruit of each person

Dim array2D(1,4)

' (column , row) --> (1,4) --> represents 2 columns and 5 rows

' Assigning values to the 2D array

array2D(0,0) = "TestUser1"
array2D(0,1) = "TestUser2"
array2D(0,2) = "TestUser3"
array2D(0,3) = "TestUser4"
array2D(0,4) = "TestUser5"

array2D(1,0) = "Apple"
array2D(1,1) = "Mango"
array2D(1,2) = "Banana"
array2D(1,3) = "Grape"
array2D(1,4) = "Berry"

MsgBox array2D(0,0)& " likes " & array2D(1,0)
MsgBox array2D(0,1)& " likes " & array2D(1,1) , 0 , SITE_TITLE
MsgBox array2D(0,2)& " likes " & array2D(1,2) , 0 , SITE_TITLE

'=============================================================
' Resizing the array :
' ReDim - we can use ReDim to resize an declared Array
' Preserve - we use Preserve keyword to preserve the existing data in an array, so in case we resize the array the data is not lost
'=============================================================

' Declaring a dynamic array 
Dim ArrayNew()

' Resizing the array to size 5
Redim ArrayNew(4)
ArrayNew(0) = "Test1"
ArrayNew(1) = "Test2"
ArrayNew(2) = "Test3"
ArrayNew(3) = "Test4"
ArrayNew(4) = "Test5"

MsgBox ArrayNew(0) , 0 , SITE_TITLE 

MsgBox ArrayNew(2) , 0 , SITE_TITLE 

' Resizing the array to size 6
' We can preserve the previously intialized array and resize the array

ReDim Preserve ArrayNew(5)

ArrayNew(5) = "Test9"

MsgBox ArrayNew(0) , 0 , SITE_TITLE 

MsgBox ArrayNew(2) , 0 , SITE_TITLE 

MsgBox ArrayNew(5) , 0 , SITE_TITLE 

