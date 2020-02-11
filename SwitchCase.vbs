' Sample program to illustrate switch cases in VB script 

Option Explicit 
On Error Resume Next 

' Declare the variable 

Dim cardType 

Cardtype = "MasterCard"

SELECT Case Cardtype
	Case "MasterCard"
		MsgBox "The card selected is MasterCard"
	case "VISA"
		MsgBox "The card selected is VISA"
	Case "American Express"
		MsgBox "The card selected is MasterCard"
	Case Else 
		MsgBox "Unknown Card"
End SELECT 