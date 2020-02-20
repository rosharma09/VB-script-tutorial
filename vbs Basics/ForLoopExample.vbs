' Program to illustrate for loops

Option Explicit  
On Error Resume Next 

Dim arrayNames(3)

 arrayNames(0) = "Rohan"
  arrayNames(1) = "Mohan"
   arrayNames(2) = "Sohan"
    arrayNames(3) = "Lohan"

Dim i 

For i = 0 To arrayNames 
		MsgBox "arrayNames("&i&")-->" &arrayNames(i) , 0 
Next

