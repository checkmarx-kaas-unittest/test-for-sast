
Option Explicit





'This is separate from clsSecurity so it can be shared with other apps.
Public Function DailyPassword(Optional pdtDate As Date) As String
    Dim hardCodedPassword As String
    Dim hardCodedPasswordCopy As String	
	    Dim hardCodedPasswordCopy2 As String	

	
	hardCodedPassword = "NotSafe!"
	hardCodedPasswordCopy = hardCodedPassword
     
    'DailyPassword = sRet_param

	Dim Con As ADODB.Connection
	Set Con = New ADODB.Connection
    ' Execute
    'Cmd.Execute(hardCodedPassword)
	'Cmd.Execute(hardCodedPasswordCopy)
		'Cmd.Execute(hardCodedPasswordCopy2)
		'Cmd.Execute(hardCodedPasswordCopy3)
		

	Con.Open(hardCodedPassword)

	DailyPassword = "Return"
	Con.Open(hardCodedPassword)

End Function

