'=====================================
' Operators in VB scripting 
' We have the following operators 
' 1. Arithematic operator : 
' 2. Comparison
' 3. Logical
'=====================================

Option Explicit 
On Error Resume Next 

' Declare variable 

Const SITE_TITLE = "www.operatorConcept.com"

Dim num1 , num2 
Dim sum , product , power 

num1 = 2
num2 = 2

sum = num1 + num2  
product = num1 * num2 
power = num1 ^ num2 

MsgBox "The sum of " &num1& " and " &num2& " is " &sum , 0 ,  SITE_TITLE

MsgBox "The product of " &num1& " and " &num2& " is " &product , 0 ,  SITE_TITLE

MsgBox num1& " power " &num2& " is " &power , 0 ,  SITE_TITLE