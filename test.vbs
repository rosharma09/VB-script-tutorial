Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to continue ?"    ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "MsgBox Demonstration"    ' Define title.
Help = "DEMO.HLP"    ' Define Help file.
Ctxt = 1000    ' Define topic context. 
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Yes"    ' Perform some action.
Else    ' User chose No.
    MyString = "No"    ' Perform some action.
End If