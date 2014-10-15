<% 
'-------------------------------------------------------------------------------
' Functions for default date format
' ANamedFormat = 0-7, where 0-4 same as VBScript
' 5 = "yyyy/mm/dd"
' 6 = "mm/dd/yyyy"
' 7 = "dd/mm/yyyy"

Function EW_FormatDateTime(ADate, ANamedFormat)
    If isdate(ADate) Then
		If ANamedFormat >= 0 and ANamedFormat <= 4 Then
			EW_FormatDateTime = FormatDateTime(ADate, ANameFormat)
		ElseIf ANamedFormat = 5 Then
			EW_FormatDateTime = Year(ADate) & "/" & Month(ADate) & "/" & Day(ADate)
		ElseIf ANamedFormat = 6 Then
			EW_FormatDateTime = Month(ADate) & "/" & Day(ADate) & "/" & Year(ADate)
		ElseIf ANamedFormat = 7 Then
			EW_FormatDateTime = Day(ADate) & "/" & Month(ADate) & "/" & Year(ADate)
		Else
			EW_FormatDateTime = ADate
		End If
	Else
		EW_FormatDateTime = ADate
    End If
End Function

Function EW_UnFormatDateTime(ADate, ANamedFormat)
	Dim arDate, AYear, AMonth, ADay
	arDate = split(ADate, "/")
	If ubound(arDate) = 2 Then
		If ANamedFormat = 5 Then
			EW_UnFormatDateTime = arDate(0) & "/" & arDate(1) & "/" & arDate(2)
		ElseIf ANamedFormat = 6 Then
			EW_UnFormatDateTime = arDate(2) & "/" & arDate(0) & "/" & arDate(1)
		ElseIf ANamedFormat = 7 Then
			EW_UnFormatDateTime = arDate(2) & "/" & arDate(1) & "/" & arDate(0)
		Else
			EW_UnFormatDateTime = ADate
		End If
	Else
		EW_UnFormatDateTime = ADate
	End If
End Function
%>

<%
'-------------------------------------------------------------------------------
' Functions for file upload

Function stringToByte(toConv)

	 For i = 1 to Len(toConv)
	 	tempChar = Mid(toConv, i, 1)
		stringToByte = stringToByte & chrB(AscB(tempChar))
	 Next
	 
End Function

Function byteToString(toConv)
	For i = 1 to LenB(toConv)
		byteord = AscB(MidB(toConv, i, 1))
		If byteord < &H80 Then ' Ascii
			byteToString = byteToString & Chr(byteord)
		Else ' Double-byte characters?
			If i < LenB(toConv) Then
				nextbyteord = AscB(MidB(toConv, i+1, 1))
				On Error Resume Next
				' Note: This line does NOT work on all systems due to limitation of the
				' Chr() function
	      byteToString = byteToString & Chr(CInt(byteord) * &H100 + CInt(nextbyteord))
				If Err.Number <> 0 Then
					On Error GoTo 0
					byteToString = byteToString & Chr(byteord) & Chr(nextbyteord)
				End If
				i = i + 1
			ElseIf i = LenB(toConv) Then
				byteToString = byteToString & Chr(byteord)
			End If
		End If
	Next
End Function

Function getValue(name)
	If dict.Exists(name) Then
		gv = CStr(dict(name).Item("Value"))	
		gv = Left(gv,Len(gv)-2)
		getValue = gv
	Else
		getValue = ""
	End If
End Function

Function getFileData(name)
	If dict.Exists(name) Then
		getFileData = dict(name).Item("Value")
		If LenB(getFileData) Mod 2 = 1 Then
			getFileData = getfileData & ChrB(0)
		End If
	Else
		getFileData = ""
	End If
End Function

Function getFileName(name)
	If dict.Exists(name) Then
		temp = dict(name).Item("FileName")
		tempPos = 1 + InStrRev(temp, "\")
		getFileName = Mid(temp, tempPos)
	Else
		getFileName = ""
	End If
End Function

Function getFileSize(name)
	If dict.Exists(name) Then
		getFileSize = LenB(dict(name).Item("Value"))
	Else
		getFileSize = 0
	End If
End Function

Function getFileContentType(name)
	If dict.Exists(name) Then
		getFileContentType = dict(name).Item("ContentType")
	Else
		getFileContentType = ""
	End If
End Function

'-------------------------------------------------------------------------------
' Note: This function does NOT work on non English servers due to limitation of
'       the Chr() function
Function saveToFile(name, path)
	If dict.Exists(name) Then
		Dim temp
			temp = dict(name).Item("Value")
		Dim fso
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
		Dim file
			Set file = fso.CreateTextFile(path)
				For tPoint = 1 to LenB(temp)
				    file.Write Chr(AscB(MidB(temp,tPoint,1)))
				Next
				file.Close
			saveToFile = True
	Else
			saveToFile = False
	End If
End Function
%>
