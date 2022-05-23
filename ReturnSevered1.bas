
Option Explicit





'This is separate from clsSecurity so it can be shared with other apps.
Public Function DailyPassword(Optional pdtDate As Date) As String
    Dim hardCodedPassword As String
    Dim hardCodedPasswordCopy As String	
	Dim hardCodedPasswordCopy2 As String	
	Dim hardCodedPasswordCopy3 As String
	Dim hardCodedPasswordCopy4 As String
	
	hardCodedPassword = "NotSafe!"
	hardCodedPasswordCopy = hardCodedPassword
     
    'DailyPassword = sRet_param

	Dim Con As ADODB.Connection
	Set Con = New ADODB.Connection
	
	Dim t
	t = InputBox()
	
	Dim Con as ADODB.connection 
    ' Execute
    Cmd.Execute(t)
	'Cmd.Execute(hardCodedPasswordCopy)
'Cmd.Execute(hardCodedPasswordCopy2)
		'Cmd.Execute(hardCodedPasswordCopy3)
		'Cmd.Execute(hardCodedPasswordCopy4)
	Con.Execute(t)

	'DailyPassword = Con.Open(hardCodedPassword)

	Con.Open(hardCodedPassword)

End Function

