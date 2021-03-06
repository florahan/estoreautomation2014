'****************************************************************
'* Function: GetStrBetween
'* Get a string from "strToSearch" between "strStart" and "strEnd"
'* Parameters: strToSearch, the full string we want to search
'* 				strStart, gap start string
'* 				strEnd, gap end string
'* For example,
'* 			GetStrBetween("<title>my title</title>", "<title>", "</title>")
'* 			return: "my title"
'****************************************************************
Function GetStrBetween(strToSearch, strStart, strEnd)
	dim startIndex, endIndex
	' Make sure the start string exists of strToSearch
	startIndex = InStr(strToSearch, strStart)
	IF NOT (startIndex > 0) THEN
		GetStrBetween = ""
		Exit Function
	END IF

	'2, make sure end string exists and after start string
	endIndex = InStr(startIndex + LEN(strStart), strToSearch, strEnd)
	IF NOT (endIndex > 0) THEN
		GetStrBetween = ""
		Exit Function
	END IF
	
	'return the between string
	GetStrBetween = MID(strToSearch, startIndex + LEN(strStart), endIndex - startIndex - LEN(strStart))
End Function


'****************************************************************
'* Function: GetStrAfter
'* Get a string after a start string, return "" if not found
'* Parameters: strToSearch, the full string we want to search
'* 				strStart, gap start string
'* For example,
'* 			GetStrAfter("user=123", "user=")
'* 			return: "123"
'****************************************************************
Function GetStrAfter(strToSearch, strStart)
	dim startIndex
	
	'1, Make sure the start string exists of strToSearch
	startIndex = InStr(strToSearch, strStart)
	IF NOT (startIndex > 0) THEN
		GetStrAfter = ""
		Exit Function
	END IF
	
	GetStrAfter = MID(strToSearch, startIndex + LEN(strStart))
END Function



'****************************************************************
'* Function: GetStrBefore
'* Get a string before a start string, return "" if not found
'* Parameters: strToSearch, the full string we want to search
'* 				strStart, gap start string
'* For example,
'* 			GetStrBefore("user=123", "=")
'* 			return: "user"
'****************************************************************
Function GetStrBefore(strToSearch, strStart)
	dim startIndex

	'1, Make sure the start string exists of strToSearch
	startIndex = InStr(strToSearch, strStart)
	IF NOT (startIndex > 0) THEN
		GetStrBefore = Null
		Exit Function
	END IF
	
	GetStrBefore = LEFT(strToSearch, startIndex - 1)
END Function


'****************************************************************
'* Function: StartsWith
'* Get a string starts with a substring or not. 
'* Return: True/False
'* Parameters: strToSearch, the full string we want to search
'* 				strStart, gap start string
'* For example,
'* 			StartsWith("user=123", "user")
'* 			return: True
'****************************************************************
Function StartsWith(strToSearch, strStart)
	IsStartsWith = False
	
	IF LEFT(strToSearch, LEN(strStart)) = strStart THEN
		StartsWith = True
	END IF
END Function



'****************************************************************
'* Function: endsWith
'* Get a string ends with a substring or not.
'* Return: True/False
'* Parameters: strToSearch, the full string we want to search
'* 				strStart, gap start string
'* For example,
'* 			endsWith("user=123", "123")
'* 			return: True
'****************************************************************
Function endsWith(strToSearch, strStart)
	endsWith = False
	IF RIGHT(strToSearch, LEN(strStart)) = strStart THEN
		endsWith = True
	END IF
END Function


'****************************************************************
'*  Function: randomStr(length)
'****************************************************************
Function randomStr(myLength)
	mySeed = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"
	mySeedLength = len(mySeed)
	
	result = ""
	Randomize
	for i = 1 to myLength
		position = Int(mySeedLength * RND + 1)
		result = result & mid(mySeed,position,1)
	next
	randomStr = result

End Function


Function randomInt(min, max)
   Randomize
   randomInt = Int((max - min + 1) * RND + min)
End Function

Function  RandSelectWebList(myWebList)
	itemCount = myWebList.GetROProperty("items count")
	'itemCount = 3
	myIndex = randomInt(0, itemCount - 1) 
 '  myIndex = randomInt(1, itemCount - 1) 
	myWebList.Select "#" & myIndex



	selectItemNames = split(myWebList.GetROProperty("selection"), ";")
	
	If selectItemNames(0) = "#0" Then
		result = ""
	else
		result = selectItemNames(0)
	End If
	RandSelectWebList = result

End Function


Function EnvironmentVarExists(myEnvVarName)
	On error resume next
	tmp = Environment(myEnvVarName)
	If Err.Number <> 0 Then
		EnvironmentVarExists = False
	else
		EnvironmentVarExists = True
	End If
	
	On error goto 0
	
End Function


Function JoinPath(firstPath, secondPath)
	if not endsWith(firstPath, "\") then
		firstPath = firstPath & "\"
	end if
	
	if StartsWith(secondPath, "\") then
		secondPath = GetStrAfter(secondPath, "\")
	end if
	
	JoinPath = firstPath & secondPath
END Function


Function DaysLater(myDayNum)
   laterDay = DateAdd("d", myDayNum, now)
End Function


Function DaysBefore(myDayNum)
   earlierDay = DateAdd("d", -myDayNum, now)
End Function

Function GetHourMinute(myDate)
   myDate = cdate(myDate)
   myHour = cstr(Hour(myDate))
   If len(myHour) = 1 Then
	   myHour = "0" & myHour
   End If

	myMinute = cstr(Minute(myDate))
	If len(myMinute) = 1 Then
		myMinute = "0" & myMinute
	End If

	GetHourMinute = myHour & ":" & myMinute
End Function



Function getReadableTime(mySecond)
	result = ""
	'get hour
	If mySecond >= 3600 Then
		myHour = mySecond \ 3600
		mySecond = mySecond - myHour * 3600 
		result = result & myHour & ":"
	End If
	
	'get minute
	myMinute = mySecond \ 60
	mySecond = mySecond - myMinute * 60
	If myMinute < 10 Then
		myMinute = "0" & myMinute
	End If
	result = result & myMinute & ":"
	
	If mySecond < 10 Then
		mySecond = "0" & mySecond
	End If
	
	result = result & mySecond
	
	getReadableTime = result
	
End Function


Function getToday(format)
	currYear = year(now)
	currMonth = month(now)
	
	if currMonth < 10 then
		currMonth = "0" & currMonth
	end if
	
	currDay = day(now)
	if currDay < 10 then
		currDay = "0" & currDay
	end if
	
	format = lcase(format)
	If format = "dd/mm/yyyy" or format = "ca" Then
		getToday = currDay & "/" & currMonth & "/" & currYear
	ElseIf format = "mm/dd//yyyy" or format = "us" Then
		getToday = currMonth  & "/" & currDay & "/" & currYear
	ElseIf format = "yyyy-mm-dd" Then
		getToday = currYear  & "-" & currMonth & "-" & currDay 
	Else
		getToday = currYear & currMonth & currDay
	End If

End Function


Function getTime(format)
	currHour = hour(now)
	if currHour < 10 then
		currHour = "0" & currHour
	end if
	
	currMinute = minute(now)
	if currMinute < 10 then
		currMinute = "0" & currMinute
	end if
	
	currSecond = second(now)
	If currSecond < 10 Then
		currSecond = "0" & currSecond
	End If
	
	getTime = currHour & format & currMinute & format & currSecond
	
End Function