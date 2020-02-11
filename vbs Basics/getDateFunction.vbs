' Prgoram to illustrate the date function with SwitchCase 

Option Explicit 
'On Error Resume Next 

' Declare a variable to store the date 

Dim CurrentDate , CurrentDay

'We use date function to get the current date 
CurrentDate = date 

MsgBox "Today's date is :" &CurrentDate 

'We use the WeekDay function to get the day : weekday takes date as an input and returns a number 

CurrentDay = WeekDay(date)

MsgBox "Today is :" &CurrentDay

' 1--> Sun
' 2--> Mod
' 3--> Tues
' 4--> wed
' 5--> thr
' 6--> Fri
' 7--> sat

SELECT Case CurrentDay
	Case "1"
		MsgBox "Sunday"
	
	Case "2"
		MsgBox "Monday"
	
	Case "3"
		MsgBox "Tuesday"
		
	Case "4"
		MsgBox "Wednesday"
		
	Case "5"
		MsgBox "Thrusday"
		
	Case "6"
		MsgBox "Friday"
	
	Case "7"
		MsgBox "Saturday"
		
End SELECT 

