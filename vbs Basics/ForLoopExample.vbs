' Program to illustrate for loops

Option Explicit  
'On Error Resume Next 

Dim a , b , myloop

a = 10
b = 1



' Program to get the sum of all the array elements of a given array 

Dim arr1 , sum , i , j 

sum = 0

arr1 = Array(10,12,1,20,3,5,2,67,40,22)

' Traverse through the elements in the array 

For i = LBound(arr1) To UBound(arr1) Step 1 
	sum = sum + arr1(i)
Next 
MsgBox "Sum of elements till now is : "  &sum

' Program to display the index of the element which is greater than the element the current element being pointed in the array, return -1 in case no such element is found

For i = LBound(arr1) To UBound(arr1) Step 1
	for j = LBound(arr1) + i To UBound(arr1) step 1 
	
	Next
Next

