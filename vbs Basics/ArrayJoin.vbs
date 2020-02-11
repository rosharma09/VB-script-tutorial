'================================
' Array join example
'================================

Option Explicit 
On Error Resume Next 

' Declare a variable array 

Dim arr(2)

arr(0) = "www"
arr(1) = "google"
arr(2) = "com"

' To join the array elements into a single string 

Dim ArrayJoin
ArrayJoin = Join(arr, ".")

MsgBox ArrayJoin