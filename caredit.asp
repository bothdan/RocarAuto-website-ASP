<% If Session("rocar_status") <> "login" Then Response.Redirect "login.asp" %>
<% Session.Timeout = 300 %>
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<% Response.Buffer = True %>
<%
If Request.TotalBytes > 0 Then
	rawData = Request.BinaryRead(Request.TotalBytes)
	separator = MidB(rawData, 1, InstrB(1, rawData, ChrB(13)) - 1)
	lenSeparator = LenB(separator)
	Set dict = Server.CreateObject("Scripting.Dictionary")
	currentPos = 1
	inStrByte = 1
	tempValue = ""	
	While inStrByte > 0
		inStrByte = InStrB(currentPos, rawData, separator)
		mValue = inStrByte - currentPos
		If mValue > 1 Then
			value = MidB(rawData, currentPos, mValue)
			Set intDict = Server.CreateObject("Scripting.Dictionary")
			begPos = 1 + InStrB(1, value, ChrB(34))
			endPos = InStrB(begPos + 1, value, ChrB(34))
			nValue = endPos
			nameN = MidB(value, begPos, endPos - begPos)
			isValid = True
			If InStrB(1, value, stringToByte("Content-Type")) > 1 Then
				begPos = 1 + InStrB(endPos + 1, value, ChrB(34))
				endPos = InStrB(begPos + 1, value, ChrB(34))
				If endPos = 0 Then
					endPos = begPos + 1
					isValid = False
				End If
				midValue = MidB(value, begPos, endPos - begPos)
				intDict.Add "FileName", Trim(byteToString(midValue))
				begPos = 14 + InStrB(endPos + 1, value, stringToByte("Content-Type:"))
				endPos = InStrB(begPos, value, ChrB(13))
				midValue = MidB(value, begPos, endPos - begPos)
				intDict.Add "ContentType", Trim(byteToString(midValue))
				begPos = endPos + 4
				endPos = LenB(value)
				nameValue = MidB(value, begPos, ((endPos - begPos) - 1))
			Else
				nameValue = Trim(byteToString(MidB(value, nValue + 5)))
			End If
			If isValid = True Then
				intDict.Add "Value", nameValue
				intDict.Add "Name", nameN
				dict.Add byteToString(nameN), intDict
			End If
		End If
		currentPos = lenSeparator + inStrByte
	Wend
	' get action
	a = getValue("a")
	key = getValue("key")
	EW_Max_File_Size = getValue("EW_Max_File_Size")
	' for the blob field
	fs_x_photo_1 = getFileSize("x_photo_1")
	' check the file size
	If fs_x_photo_1 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_1 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_1 = getFileName("x_photo_1")
	ct_x_photo_1 = getFileContentType("x_photo_1")
	x_photo_1 = getFileData("x_photo_1")
	w_x_photo_1 = getValue("w_x_photo_1")
	h_x_photo_1 = getValue("h_x_photo_1")
	a_x_photo_1 = getValue("a_x_photo_1")
	' for the blob field
	fs_x_photo_2 = getFileSize("x_photo_2")
	' check the file size
	If fs_x_photo_2 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_2 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_2 = getFileName("x_photo_2")
	ct_x_photo_2 = getFileContentType("x_photo_2")
	x_photo_2 = getFileData("x_photo_2")
	w_x_photo_2 = getValue("w_x_photo_2")
	h_x_photo_2 = getValue("h_x_photo_2")
	a_x_photo_2 = getValue("a_x_photo_2")
	' for the blob field
	fs_x_photo_3 = getFileSize("x_photo_3")
	' check the file size
	If fs_x_photo_3 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_3 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_3 = getFileName("x_photo_3")
	ct_x_photo_3 = getFileContentType("x_photo_3")
	x_photo_3 = getFileData("x_photo_3")
	w_x_photo_3 = getValue("w_x_photo_3")
	h_x_photo_3 = getValue("h_x_photo_3")
	a_x_photo_3 = getValue("a_x_photo_3")
	' for the blob field
	fs_x_photo_4 = getFileSize("x_photo_4")
	' check the file size
	If fs_x_photo_4 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_4 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_4 = getFileName("x_photo_4")
	ct_x_photo_4 = getFileContentType("x_photo_4")
	x_photo_4 = getFileData("x_photo_4")
	w_x_photo_4 = getValue("w_x_photo_4")
	h_x_photo_4 = getValue("h_x_photo_4")
	a_x_photo_4 = getValue("a_x_photo_4")
	' for the blob field
	fs_x_photo_5 = getFileSize("x_photo_5")
	' check the file size
	If fs_x_photo_5 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_5 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_5 = getFileName("x_photo_5")
	ct_x_photo_5 = getFileContentType("x_photo_5")
	x_photo_5 = getFileData("x_photo_5")
	w_x_photo_5 = getValue("w_x_photo_5")
	h_x_photo_5 = getValue("h_x_photo_5")
	a_x_photo_5 = getValue("a_x_photo_5")
	' for the blob field
	fs_x_photo_6 = getFileSize("x_photo_6")
	' check the file size
	If fs_x_photo_6 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_6 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_6 = getFileName("x_photo_6")
	ct_x_photo_6 = getFileContentType("x_photo_6")
	x_photo_6 = getFileData("x_photo_6")
	w_x_photo_6 = getValue("w_x_photo_6")
	h_x_photo_6 = getValue("h_x_photo_6")
	a_x_photo_6 = getValue("a_x_photo_6")
	' for the blob field
	fs_x_photo_7 = getFileSize("x_photo_7")
	' check the file size
	If fs_x_photo_7 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_7 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_7 = getFileName("x_photo_7")
	ct_x_photo_7 = getFileContentType("x_photo_7")
	x_photo_7 = getFileData("x_photo_7")
	w_x_photo_7 = getValue("w_x_photo_7")
	h_x_photo_7 = getValue("h_x_photo_7")
	a_x_photo_7 = getValue("a_x_photo_7")
	' for the blob field
	fs_x_photo_8 = getFileSize("x_photo_8")
	' check the file size
	If fs_x_photo_8 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_8 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_8 = getFileName("x_photo_8")
	ct_x_photo_8 = getFileContentType("x_photo_8")
	x_photo_8 = getFileData("x_photo_8")
	w_x_photo_8 = getValue("w_x_photo_8")
	h_x_photo_8 = getValue("h_x_photo_8")
	a_x_photo_8 = getValue("a_x_photo_8")
	' for the blob field
	fs_x_photo_9 = getFileSize("x_photo_9")
	' check the file size
	If fs_x_photo_9 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_9 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_9 = getFileName("x_photo_9")
	ct_x_photo_9 = getFileContentType("x_photo_9")
	x_photo_9 = getFileData("x_photo_9")
	w_x_photo_9 = getValue("w_x_photo_9")
	h_x_photo_9 = getValue("h_x_photo_9")
	a_x_photo_9 = getValue("a_x_photo_9")
	' for the blob field
	fs_x_photo_10 = getFileSize("x_photo_10")
	' check the file size
	If fs_x_photo_10 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_10 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_10 = getFileName("x_photo_10")
	ct_x_photo_10 = getFileContentType("x_photo_10")
	x_photo_10 = getFileData("x_photo_10")
	w_x_photo_10 = getValue("w_x_photo_10")
	h_x_photo_10 = getValue("h_x_photo_10")
	a_x_photo_10 = getValue("a_x_photo_10")
	' for the blob field
	fs_x_photo_11 = getFileSize("x_photo_11")
	' check the file size
	If fs_x_photo_11 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_11 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_11 = getFileName("x_photo_11")
	ct_x_photo_11 = getFileContentType("x_photo_11")
	x_photo_11 = getFileData("x_photo_11")
	w_x_photo_11 = getValue("w_x_photo_11")
	h_x_photo_11 = getValue("h_x_photo_11")
	a_x_photo_11 = getValue("a_x_photo_11")
	' for the blob field
	fs_x_photo_12 = getFileSize("x_photo_12")
	' check the file size
	If fs_x_photo_12 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_12 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_12 = getFileName("x_photo_12")
	ct_x_photo_12 = getFileContentType("x_photo_12")
	x_photo_12 = getFileData("x_photo_12")
	w_x_photo_12 = getValue("w_x_photo_12")
	h_x_photo_12 = getValue("h_x_photo_12")
	a_x_photo_12 = getValue("a_x_photo_12")
	' for the blob field
	fs_x_photo_13 = getFileSize("x_photo_13")
	' check the file size
	If fs_x_photo_13 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_13 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_13 = getFileName("x_photo_13")
	ct_x_photo_13 = getFileContentType("x_photo_13")
	x_photo_13 = getFileData("x_photo_13")
	w_x_photo_13 = getValue("w_x_photo_13")
	h_x_photo_13 = getValue("h_x_photo_13")
	a_x_photo_13 = getValue("a_x_photo_13")
	' for the blob field
	fs_x_photo_14 = getFileSize("x_photo_14")
	' check the file size
	If fs_x_photo_14 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_14 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_14 = getFileName("x_photo_14")
	ct_x_photo_14 = getFileContentType("x_photo_14")
	x_photo_14 = getFileData("x_photo_14")
	w_x_photo_14 = getValue("w_x_photo_14")
	h_x_photo_14 = getValue("h_x_photo_14")
	a_x_photo_14 = getValue("a_x_photo_14")
	' for the blob field
	fs_x_photo_15 = getFileSize("x_photo_15")
	' check the file size
	If fs_x_photo_15 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_15 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_15 = getFileName("x_photo_15")
	ct_x_photo_15 = getFileContentType("x_photo_15")
	x_photo_15 = getFileData("x_photo_15")
	w_x_photo_15 = getValue("w_x_photo_15")
	h_x_photo_15 = getValue("h_x_photo_15")
	a_x_photo_15 = getValue("a_x_photo_15")
	' for the blob field
	fs_x_photo_16 = getFileSize("x_photo_16")
	' check the file size
	If fs_x_photo_16 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_16 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_16 = getFileName("x_photo_16")
	ct_x_photo_16 = getFileContentType("x_photo_16")
	x_photo_16 = getFileData("x_photo_16")
	w_x_photo_16 = getValue("w_x_photo_16")
	h_x_photo_16 = getValue("h_x_photo_16")
	a_x_photo_16 = getValue("a_x_photo_16")
	' for the blob field
	fs_x_photo_17 = getFileSize("x_photo_17")
	' check the file size
	If fs_x_photo_17 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_17 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_17 = getFileName("x_photo_17")
	ct_x_photo_17 = getFileContentType("x_photo_17")
	x_photo_17 = getFileData("x_photo_17")
	w_x_photo_17 = getValue("w_x_photo_17")
	h_x_photo_17 = getValue("h_x_photo_17")
	a_x_photo_17 = getValue("a_x_photo_17")
	' for the blob field
	fs_x_photo_18 = getFileSize("x_photo_18")
	' check the file size
	If fs_x_photo_18 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_18 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_18 = getFileName("x_photo_18")
	ct_x_photo_18 = getFileContentType("x_photo_18")
	x_photo_18 = getFileData("x_photo_18")
	w_x_photo_18 = getValue("w_x_photo_18")
	h_x_photo_18 = getValue("h_x_photo_18")
	a_x_photo_18 = getValue("a_x_photo_18")
	' for the blob field
	fs_x_photo_19 = getFileSize("x_photo_19")
	' check the file size
	If fs_x_photo_19 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_19 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_19 = getFileName("x_photo_19")
	ct_x_photo_19 = getFileContentType("x_photo_19")
	x_photo_19 = getFileData("x_photo_19")
	w_x_photo_19 = getValue("w_x_photo_19")
	h_x_photo_19 = getValue("h_x_photo_19")
	a_x_photo_19 = getValue("a_x_photo_19")
	' for the blob field
	fs_x_photo_20 = getFileSize("x_photo_20")
	' check the file size
	If fs_x_photo_20 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_20 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_20 = getFileName("x_photo_20")
	ct_x_photo_20 = getFileContentType("x_photo_20")
	x_photo_20 = getFileData("x_photo_20")
	w_x_photo_20 = getValue("w_x_photo_20")
	h_x_photo_20 = getValue("h_x_photo_20")
	a_x_photo_20 = getValue("a_x_photo_20")
	' for the blob field
	fs_x_photo_21 = getFileSize("x_photo_21")
	' check the file size
	If fs_x_photo_21 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_21 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_21 = getFileName("x_photo_21")
	ct_x_photo_21 = getFileContentType("x_photo_21")
	x_photo_21 = getFileData("x_photo_21")
	w_x_photo_21 = getValue("w_x_photo_21")
	h_x_photo_21 = getValue("h_x_photo_21")
	a_x_photo_21 = getValue("a_x_photo_21")
	' for the blob field
	fs_x_photo_22 = getFileSize("x_photo_22")
	' check the file size
	If fs_x_photo_22 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_22 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_22 = getFileName("x_photo_22")
	ct_x_photo_22 = getFileContentType("x_photo_22")
	x_photo_22 = getFileData("x_photo_22")
	w_x_photo_22 = getValue("w_x_photo_22")
	h_x_photo_22 = getValue("h_x_photo_22")
	a_x_photo_22 = getValue("a_x_photo_22")
	' for the blob field
	fs_x_photo_23 = getFileSize("x_photo_23")
	' check the file size
	If fs_x_photo_23 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_23 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_23 = getFileName("x_photo_23")
	ct_x_photo_23 = getFileContentType("x_photo_23")
	x_photo_23 = getFileData("x_photo_23")
	w_x_photo_23 = getValue("w_x_photo_23")
	h_x_photo_23 = getValue("h_x_photo_23")
	a_x_photo_23 = getValue("a_x_photo_23")
	' for the blob field
	fs_x_photo_24 = getFileSize("x_photo_24")
	' check the file size
	If fs_x_photo_24 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_24 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_24 = getFileName("x_photo_24")
	ct_x_photo_24 = getFileContentType("x_photo_24")
	x_photo_24 = getFileData("x_photo_24")
	w_x_photo_24 = getValue("w_x_photo_24")
	h_x_photo_24 = getValue("h_x_photo_24")
	a_x_photo_24 = getValue("a_x_photo_24")
	' for the blob field
	fs_x_photo_25 = getFileSize("x_photo_25")
	' check the file size
	If fs_x_photo_25 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_25 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_25 = getFileName("x_photo_25")
	ct_x_photo_25 = getFileContentType("x_photo_25")
	x_photo_25 = getFileData("x_photo_25")
	w_x_photo_25 = getValue("w_x_photo_25")
	h_x_photo_25 = getValue("h_x_photo_25")
	a_x_photo_25 = getValue("a_x_photo_25")
	' for the blob field
	fs_x_photo_26 = getFileSize("x_photo_26")
	' check the file size
	If fs_x_photo_26 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_26 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_26 = getFileName("x_photo_26")
	ct_x_photo_26 = getFileContentType("x_photo_26")
	x_photo_26 = getFileData("x_photo_26")
	w_x_photo_26 = getValue("w_x_photo_26")
	h_x_photo_26 = getValue("h_x_photo_26")
	a_x_photo_26 = getValue("a_x_photo_26")
	' for the blob field
	fs_x_photo_27 = getFileSize("x_photo_27")
	' check the file size
	If fs_x_photo_27 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_27 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_27 = getFileName("x_photo_27")
	ct_x_photo_27 = getFileContentType("x_photo_27")
	x_photo_27 = getFileData("x_photo_27")
	w_x_photo_27 = getValue("w_x_photo_27")
	h_x_photo_27 = getValue("h_x_photo_27")
	a_x_photo_27 = getValue("a_x_photo_27")
	' for the blob field
	fs_x_photo_28 = getFileSize("x_photo_28")
	' check the file size
	If fs_x_photo_28 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_28 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_28 = getFileName("x_photo_28")
	ct_x_photo_28 = getFileContentType("x_photo_28")
	x_photo_28 = getFileData("x_photo_28")
	w_x_photo_28 = getValue("w_x_photo_28")
	h_x_photo_28 = getValue("h_x_photo_28")
	a_x_photo_28 = getValue("a_x_photo_28")
	' for the blob field
	fs_x_photo_29 = getFileSize("x_photo_29")
	' check the file size
	If fs_x_photo_29 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_29 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_29 = getFileName("x_photo_29")
	ct_x_photo_29 = getFileContentType("x_photo_29")
	x_photo_29 = getFileData("x_photo_29")
	w_x_photo_29 = getValue("w_x_photo_29")
	h_x_photo_29 = getValue("h_x_photo_29")
	a_x_photo_29 = getValue("a_x_photo_29")
	' for the blob field
	fs_x_photo_30 = getFileSize("x_photo_30")
	' check the file size
	If fs_x_photo_30 > 0 And CLng(EW_Max_File_Size) > 0 Then
		If fs_x_photo_30 > CLng(EW_Max_File_Size) Then
			Response.Write "Max. file size (" & EW_Max_File_Size & " bytes) exceeded."
			Response.End
		End If
	End If	
	fn_x_photo_30 = getFileName("x_photo_30")
	ct_x_photo_30 = getFileContentType("x_photo_30")
	x_photo_30 = getFileData("x_photo_30")
	w_x_photo_30 = getValue("w_x_photo_30")
	h_x_photo_30 = getValue("h_x_photo_30")
	a_x_photo_30 = getValue("a_x_photo_30")
	' for other fields
	x_ID = getValue("x_ID")
	x_year = getValue("x_year")
	x_make = getValue("x_make")
	x_model = getValue("x_model")
	x_type = getValue("x_type")
	x_miles = getValue("x_miles")
	x_price = getValue("x_price")
	x_doors = getValue("x_doors")
	x_engine = getValue("x_engine")
	x_transmission = getValue("x_transmission")
	x_drivetrain = getValue("x_drivetrain")
	x_ext_color = getValue("x_ext_color")
	x_int_color = getValue("x_int_color")
	x_stock = getValue("x_stock")
	x_vin = getValue("x_vin")
	x_city_mpg = getValue("x_city_mpg")
	x_hwy_mpg = getValue("x_hwy_mpg")
	x_carfax = getValue("x_carfax")
	x_special = getValue("x_special")
	x_status = getValue("x_status")
	x_features = getValue("x_features")
	If IsObject(intDict) Then
		intDict.RemoveAll
		Set intDict = Nothing
	End If
	dict.RemoveAll
	Set dict = Nothing
Else
	key = Request.Querystring("key")
	a = "I" 'display record
End If
If key="" OR IsNull(key) Then Response.Redirect "carlist.asp"
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [car] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "carlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
			x_ID = rs("ID")
			x_year = rs("year")
			x_make = rs("make")
			x_model = rs("model")
			x_type = rs("type")
			x_miles = rs("miles")
			x_price = rs("price")
			x_doors = rs("doors")
			x_engine = rs("engine")
			x_transmission = rs("transmission")
			x_drivetrain = rs("drivetrain")
			x_ext_color = rs("ext_color")
			x_int_color = rs("int_color")
			x_stock = rs("stock")
			x_vin = rs("vin")
			x_city_mpg = rs("city_mpg")
			x_hwy_mpg = rs("hwy_mpg")
			x_carfax = rs("carfax")
			x_special = rs("special")
			x_status = rs("status")
			x_features = rs("features")
			x_photo_1 = rs("photo 1")
			x_photo_2 = rs("photo 2")
			x_photo_3 = rs("photo 3")
			x_photo_4 = rs("photo 4")
			x_photo_5 = rs("photo 5")
			x_photo_6 = rs("photo 6")
			x_photo_7 = rs("photo 7")
			x_photo_8 = rs("photo 8")
			x_photo_9 = rs("photo 9")
			x_photo_10 = rs("photo 10")
			x_photo_11 = rs("photo 11")
			x_photo_12 = rs("photo 12")
			x_photo_13 = rs("photo 13")
			x_photo_14 = rs("photo 14")
			x_photo_15 = rs("photo 15")
			x_photo_16 = rs("photo 16")
			x_photo_17 = rs("photo 17")
			x_photo_18 = rs("photo 18")
			x_photo_19 = rs("photo 19")
			x_photo_20 = rs("photo 20")
			x_photo_21 = rs("photo 21")
			x_photo_22 = rs("photo 22")
			x_photo_23 = rs("photo 23")
			x_photo_24 = rs("photo 24")
			x_photo_25 = rs("photo 25")
			x_photo_26 = rs("photo 26")
			x_photo_27 = rs("photo 27")
			x_photo_28 = rs("photo 28")
			x_photo_29 = rs("photo 29")
			x_photo_30 = rs("photo 30")
		End If
		rs.Close
		Set rs = Nothing
	Case "U": ' Update
		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [car] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.EOF Then
			Response.Clear
			Response.Redirect "carlist.asp"
		End If
		tmpFld = Trim(x_year)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("year") = tmpFld
		tmpFld = Trim(x_make)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("make") = tmpFld
		tmpFld = Trim(x_model)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("model") = tmpFld
		tmpFld = Trim(x_type)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("type") = tmpFld
		tmpFld = x_miles
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("miles") = cLng(tmpFld)
		tmpFld = x_price
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("price") = cDbl(tmpFld)
		tmpFld = x_doors
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("doors") = cLng(tmpFld)
		tmpFld = Trim(x_engine)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("engine") = tmpFld
		tmpFld = Trim(x_transmission)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("transmission") = tmpFld
		tmpFld = Trim(x_drivetrain)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("drivetrain") = tmpFld
		tmpFld = Trim(x_ext_color)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("ext_color") = tmpFld
		tmpFld = Trim(x_int_color)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("int_color") = tmpFld
		tmpFld = Trim(x_stock)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("stock") = tmpFld
		tmpFld = Trim(x_vin)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("vin") = tmpFld
		tmpFld = Trim(x_city_mpg)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("city_mpg") = tmpFld
		tmpFld = Trim(x_hwy_mpg)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("hwy_mpg") = tmpFld
		tmpFld = Trim(x_carfax)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("carfax") = tmpFld
		tmpFld = Trim(x_special)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("special") = tmpFld
		
		tmpFld = Trim(x_status)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("status") = tmpFld
		
		tmpFld = Trim(x_features)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("features") = tmpFld
		
		If a_x_photo_1 = "2" Or a_x_photo_1 = "3" Then 'Update/Remove
		tmpFld = x_photo_1
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 1") = Null
		Else
		rs("photo 1").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_2 = "2" Or a_x_photo_2 = "3" Then 'Update/Remove
		tmpFld = x_photo_2
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 2") = Null
		Else
		rs("photo 2").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_3 = "2" Or a_x_photo_3 = "3" Then 'Update/Remove
		tmpFld = x_photo_3
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 3") = Null
		Else
		rs("photo 3").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_4 = "2" Or a_x_photo_4 = "3" Then 'Update/Remove
		tmpFld = x_photo_4
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 4") = Null
		Else
		rs("photo 4").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_5 = "2" Or a_x_photo_5 = "3" Then 'Update/Remove
		tmpFld = x_photo_5
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 5") = Null
		Else
		rs("photo 5").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_6 = "2" Or a_x_photo_6 = "3" Then 'Update/Remove
		tmpFld = x_photo_6
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 6") = Null
		Else
		rs("photo 6").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_7 = "2" Or a_x_photo_7 = "3" Then 'Update/Remove
		tmpFld = x_photo_7
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 7") = Null
		Else
		rs("photo 7").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_8 = "2" Or a_x_photo_8 = "3" Then 'Update/Remove
		tmpFld = x_photo_8
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 8") = Null
		Else
		rs("photo 8").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_9 = "2" Or a_x_photo_9 = "3" Then 'Update/Remove
		tmpFld = x_photo_9
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 9") = Null
		Else
		rs("photo 9").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_10 = "2" Or a_x_photo_10 = "3" Then 'Update/Remove
		tmpFld = x_photo_10
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 10") = Null
		Else
		rs("photo 10").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_11 = "2" Or a_x_photo_11 = "3" Then 'Update/Remove
		tmpFld = x_photo_11
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 11") = Null
		Else
		rs("photo 11").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_12 = "2" Or a_x_photo_12 = "3" Then 'Update/Remove
		tmpFld = x_photo_12
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 12") = Null
		Else
		rs("photo 12").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_13 = "2" Or a_x_photo_13 = "3" Then 'Update/Remove
		tmpFld = x_photo_13
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 13") = Null
		Else
		rs("photo 13").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_14 = "2" Or a_x_photo_14 = "3" Then 'Update/Remove
		tmpFld = x_photo_14
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 14") = Null
		Else
		rs("photo 14").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_15 = "2" Or a_x_photo_15 = "3" Then 'Update/Remove
		tmpFld = x_photo_15
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 15") = Null
		Else
		rs("photo 15").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_16 = "2" Or a_x_photo_16 = "3" Then 'Update/Remove
		tmpFld = x_photo_16
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 16") = Null
		Else
		rs("photo 16").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_17 = "2" Or a_x_photo_17 = "3" Then 'Update/Remove
		tmpFld = x_photo_17
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 17") = Null
		Else
		rs("photo 17").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_18 = "2" Or a_x_photo_18 = "3" Then 'Update/Remove
		tmpFld = x_photo_18
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 18") = Null
		Else
		rs("photo 18").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_19 = "2" Or a_x_photo_19 = "3" Then 'Update/Remove
		tmpFld = x_photo_19
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 19") = Null
		Else
		rs("photo 19").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_20 = "2" Or a_x_photo_20 = "3" Then 'Update/Remove
		tmpFld = x_photo_20
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 20") = Null
		Else
		rs("photo 20").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_21 = "2" Or a_x_photo_21 = "3" Then 'Update/Remove
		tmpFld = x_photo_21
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 21") = Null
		Else
		rs("photo 21").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_22 = "2" Or a_x_photo_22 = "3" Then 'Update/Remove
		tmpFld = x_photo_22
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 22") = Null
		Else
		rs("photo 22").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_23 = "2" Or a_x_photo_23 = "3" Then 'Update/Remove
		tmpFld = x_photo_23
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 23") = Null
		Else
		rs("photo 23").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_24 = "2" Or a_x_photo_24 = "3" Then 'Update/Remove
		tmpFld = x_photo_24
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 24") = Null
		Else
		rs("photo 24").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_25 = "2" Or a_x_photo_25 = "3" Then 'Update/Remove
		tmpFld = x_photo_25
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 25") = Null
		Else
		rs("photo 25").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_26 = "2" Or a_x_photo_26 = "3" Then 'Update/Remove
		tmpFld = x_photo_26
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 26") = Null
		Else
		rs("photo 26").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_27 = "2" Or a_x_photo_27 = "3" Then 'Update/Remove
		tmpFld = x_photo_27
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 27") = Null
		Else
		rs("photo 27").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_28 = "2" Or a_x_photo_28 = "3" Then 'Update/Remove
		tmpFld = x_photo_28
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 28") = Null
		Else
		rs("photo 28").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_29 = "2" Or a_x_photo_29 = "3" Then 'Update/Remove
		tmpFld = x_photo_29
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 29") = Null
		Else
		rs("photo 29").AppendChunk tmpFld
		End If
		End If
		If a_x_photo_30 = "2" Or a_x_photo_30 = "3" Then 'Update/Remove
		tmpFld = x_photo_30
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		If IsNull(tmpFld) Then
		rs("photo 30") = Null
		Else
		rs("photo 30").AppendChunk tmpFld
		End If
		End If
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "adminlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="802" bgcolor="white">
    <tr>
        <td bgcolor="white"><script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_miles && !EW_checkinteger(EW_this.x_miles.value)) {
        if (!EW_onError(EW_this, EW_this.x_miles, "TEXT", "Incorrect integer - Miles"))
            return false; 
        }
if (EW_this.x_price && !EW_checknumber(EW_this.x_price.value)) {
        if (!EW_onError(EW_this, EW_this.x_price, "TEXT", "Incorrect floating point number - Price"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="caredit.asp" method="post" enctype="multipart/form-data">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<input type="hidden" name="EW_Max_File_Size" value="9000000">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" align="center" width="624">
<tr>
<td bgcolor="white" width="156"><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156"><font face="Arial" size="2" color="whitesmoke"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font><font color="whitesmoke">&nbsp;</font></td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Year</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_yearList = "<SELECT name='x_year'><OPTION value=''>Please Select</OPTION>"
    x_yearList = x_yearList & "<OPTION value=""2010"""
    If x_year = "2010" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2010" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2009"""
    If x_year = "2009" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2009" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2008"""
    If x_year = "2008" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2008" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2007"""
    If x_year = "2007" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2007" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2006"""
    If x_year = "2006" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2006" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2005"""
    If x_year = "2005" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2005" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2004"""
    If x_year = "2004" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2004" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2003"""
    If x_year = "2003" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2003" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2002"""
    If x_year = "2002" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2002" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2001"""
    If x_year = "2001" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2001" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2000"""
    If x_year = "2000" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2000" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1999"""
    If x_year = "1999" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1999" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1998"""
    If x_year = "1998" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1998" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1997"""
    If x_year = "1997" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1997" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1996"""
    If x_year = "1996" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1996" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1995"""
    If x_year = "1995" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1995" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1994"""
    If x_year = "1994" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1994" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1993"""
    If x_year = "1993" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1993" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1992"""
    If x_year = "1992" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1992" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1991"""
    If x_year = "1991" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1991" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1990"""
    If x_year = "1990" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1990" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1989"""
    If x_year = "1989" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1989" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1988"""
    If x_year = "1988" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1988" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1887"""
    If x_year = "1887" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1987" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1886"""
    If x_year = "1886" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1986" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1885"""
    If x_year = "1885" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1985" & "</option>"
x_yearList = x_yearList & "</select>"
response.write x_yearList
%>
</font>&nbsp;</td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Exterior color</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_ext_color" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_ext_color&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Make</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_makeList = "<SELECT name='x_make'><OPTION value=''>Please Select</OPTION>"
    x_makeList = x_makeList & "<OPTION value=""Acura"""
    If x_make = "Acura" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Acura" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Alfa Romeo"""
    If x_make = "Alfa Romeo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Alfa Romeo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Am General"""
    If x_make = "Am General" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Am General" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Aston Martin"""
    If x_make = "Aston Martin" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Aston Martin" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Audi"""
    If x_make = "Audi" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Audi" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""BMW"""
    If x_make = "BMW" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "BMW" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Bentley"""
    If x_make = "Bentley" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Bentley" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Buick"""
    If x_make = "Buick" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Buick" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Cadillac"""
    If x_make = "Cadillac" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Cadillac" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Chevrolet"""
    If x_make = "Chevrolet" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Chevrolet" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Chrysler"""
    If x_make = "Chrysler" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Chrysler" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Dacia"""
    If x_make = "Dacia" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Dacia" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Daewoo"""
    If x_make = "Daewoo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Daewoo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Daihatsu"""
    If x_make = "Daihatsu" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Daihatsu" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Dodge"""
    If x_make = "Dodge" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Dodge" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Eagle"""
    If x_make = "Eagle" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Eagle" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Ferrari"""
    If x_make = "Ferrari" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Ferrari" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Ford"""
    If x_make = "Ford" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Ford" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""GMC"""
    If x_make = "GMC" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "GMC" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Geo"""
    If x_make = "Geo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Geo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Honda"""
    If x_make = "Honda" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Honda" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Hummer"""
    If x_make = "Hummer" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Hummer" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Hyundai"""
    If x_make = "Hyundai" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Hyundai" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Infiniti"""
    If x_make = "Infiniti" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Infiniti" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""International"""
    If x_make = "International" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "International" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Isuzu"""
    If x_make = "Isuzu" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Isuzu" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Jaguar"""
    If x_make = "Jaguar" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Jaguar" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Jeep"""
    If x_make = "Jeep" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Jeep" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Kia"""
    If x_make = "Kia" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Kia" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lamborghini"""
    If x_make = "Lamborghini" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lamborghini" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Land Rover"""
    If x_make = "Land Rover" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Land Rover" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lexus"""
    If x_make = "Lexus" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lexus" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lincoln"""
    If x_make = "Lincoln" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lincoln" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lotus"""
    If x_make = "Lotus" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lotus" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Maserati"""
    If x_make = "Maserati" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Maserati" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Maybach"""
    If x_make = "Maybach" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Maybach" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mazda"""
    If x_make = "Mazda" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mazda" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mercedes-Benz"""
    If x_make = "Mercedes-Benz" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mercedes-Benz" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mercury"""
    If x_make = "Mercury" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mercury" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mini"""
    If x_make = "Mini" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mini" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mitsubishi"""
    If x_make = "Mitsubishi" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mitsubishi" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Morgan"""
    If x_make = "Morgan" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Morgan" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Nissan"""
    If x_make = "Nissan" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Nissan" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Oldsmobile"""
    If x_make = "Oldsmobile" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Oldsmobile" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Panoz"""
    If x_make = "Panoz" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Panoz" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Peugeot"""
    If x_make = "Peugeot" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Peugeot" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Plymouth"""
    If x_make = "Plymouth" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Plymouth" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Pontiac"""
    If x_make = "Pontiac" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Pontiac" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Porsche"""
    If x_make = "Porsche" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Porche" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Rolls-Royce"""
    If x_make = "Rolls-Royce" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Rolls-Royce" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Saab"""
    If x_make = "Saab" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Saab" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Saleen"""
    If x_make = "Saleen" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Saleen" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Saturn"""
    If x_make = "Saturn" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Saturn" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Scion"""
    If x_make = "Scion" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Scion" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Smart"""
    If x_make = "Smart" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Smart" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Sterling"""
    If x_make = "Sterling" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Sterling" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Subaru"""
    If x_make = "Subaru" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Subaru" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Suzuki"""
    If x_make = "Suzuki" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Suzuki" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Tesla"""
    If x_make = "Tesla" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Tesla" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Toyota"""
    If x_make = "Toyota" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Toyota" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Volkswagen"""
    If x_make = "Volkswagen" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Volkswagen" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Volvo"""
    If x_make = "Volvo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Volvo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Yugo"""
    If x_make = "Yugo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Yugo" & "</option>"
x_makeList = x_makeList & "</select>"
response.write x_makeList
%>
</font>&nbsp;</td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Interior color</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_int_color" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_int_color&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Model</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_model" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_model&"") %>"></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Stock #</span></b></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_stock" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_stock&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Engine</span></b></font></td>
<td bgcolor="white" width="156" height="25">
<font face="Arial" size="2"><input type="text" name="x_engine" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_engine&"") %>"></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">VIN</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_vin" size="20" maxlength=17 value="<%= Server.HtmlEncode(x_vin&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Miles</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_miles" value="<%= Server.HtmlEncode(x_miles&"") %>" size="20"></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">City MPG</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_city_mpg" size="20" maxlength=2 value="<%= Server.HtmlEncode(x_city_mpg&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Price</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_price" value="<%= Server.HtmlEncode(x_price&"") %>" size="20"></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Hwy MPG</span></b></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><input type="text" name="x_hwy_mpg" size="20" maxlength=2 value="<%= Server.HtmlEncode(x_hwy_mpg&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Doors</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_doorsList = "<SELECT name='x_doors'><OPTION value=''>Please Select</OPTION>"
    x_doorsList = x_doorsList & "<OPTION value=""5"""
    If x_doors = "5" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "5" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""4"""
    If x_doors = "4" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "4" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""3"""
    If x_doors = "3" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "3" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""2"""
    If x_doors = "2" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "2" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""1"""
    If x_doors = "1" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "1" & "</option>"
x_doorsList = x_doorsList & "</select>"
response.write x_doorsList
%>
</font>&nbsp;</td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Carfax</span></b></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_carfaxList = "<SELECT name='x_carfax'><OPTION value=''>Please Select</OPTION>"
    x_carfaxList = x_carfaxList & "<OPTION value=""Yes"""
    If x_carfax = "Yes" Then
        x_carfaxList = x_carfaxList & " selected"
    End If
    x_carfaxList = x_carfaxList & ">" & "Yes" & "</option>"
    x_carfaxList = x_carfaxList & "<OPTION value=""No"""
    If x_carfax = "No" Then
        x_carfaxList = x_carfaxList & " selected"
    End If
    x_carfaxList = x_carfaxList & ">" & "No" & "</option>"
x_carfaxList = x_carfaxList & "</select>"
response.write x_carfaxList
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Type</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_typeList = "<SELECT name='x_type'><OPTION value=''>Please Select</OPTION>"
    x_typeList = x_typeList & "<OPTION value=""Sedan"""
    If x_type = "Sedan" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Sedan" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""SUV"""
    If x_type = "SUV" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "SUV" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Mini-Van"""
    If x_type = "Mini-Van" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Mini-Van" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Wagon"""
    If x_type = "Wagon" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Wagon" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Hatchback"""
    If x_type = "Hatchback" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Hatchback" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Coupe"""
    If x_type = "Coupe" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Coupe" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Truck"""
    If x_type = "Truck" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Truck" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Convertible"""
    If x_type = "Convertible" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Convertible" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Sport"""
    If x_type = "Sport" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Sport" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""SUT"""
    If x_type = "SUT" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "SUT" & "</option>"
x_typeList = x_typeList & "</select>"
response.write x_typeList
%></font></td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Special</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_specialList = "<SELECT name='x_special'><OPTION value=''>Please Select</OPTION>"
    x_specialList = x_specialList & "<OPTION value=""Yes"""
    If x_special = "Yes" Then
        x_specialList = x_specialList & " selected"
    End If
    x_specialList = x_specialList & ">" & "Yes" & "</option>"
    x_specialList = x_specialList & "<OPTION value=""No"""
    If x_special = "No" Then
        x_specialList = x_specialList & " selected"
    End If
    x_specialList = x_specialList & ">" & "No" & "</option>"
x_specialList = x_specialList & "</select>"
response.write x_specialList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Transmission</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_transmissionList = "<SELECT name='x_transmission'><OPTION value=''>Please Select</OPTION>"
    x_transmissionList = x_transmissionList & "<OPTION value=""Automatic"""
    If x_transmission = "Automatic" Then
        x_transmissionList = x_transmissionList & " selected"
    End If
    x_transmissionList = x_transmissionList & ">" & "Automatic" & "</option>"
    x_transmissionList = x_transmissionList & "<OPTION value=""Manual"""
    If x_transmission = "Manual" Then
        x_transmissionList = x_transmissionList & " selected"
    End If
    x_transmissionList = x_transmissionList & ">" & "Manual" & "</option>"
x_transmissionList = x_transmissionList & "</select>"
response.write x_transmissionList
%>
</font>&nbsp;</td>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Status</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_statusList = "<SELECT name='x_status'><OPTION value=''>Please Select</OPTION>"
    x_statusList = x_statusList & "<OPTION value=""For Sale"""
    If x_status = "For Sale" Then
        x_statusList = x_statusList & " selected"
    End If
    x_statusList = x_statusList & ">" & "For Sale" & "</option>"
    x_statusList = x_statusList & "<OPTION value=""Sold"""
    If x_status = "Sold" Then
        x_statusList = x_statusList & " selected"
    End If
    x_statusList = x_statusList & ">" & "Sold" & "</option>"
x_statusList = x_statusList & "</select>"
response.write x_statusList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="156" height="25"><font face="Arial"><b><span style="font-size:10pt;">Drivetrain</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156" height="25"><font face="Arial" size="2"><%
x_drivetrainList = "<SELECT name='x_drivetrain'><OPTION value=''>Please Select</OPTION>"
    x_drivetrainList = x_drivetrainList & "<OPTION value=""FWD"""
    If x_drivetrain = "FWD" Then
        x_drivetrainList = x_drivetrainList & " selected"
    End If
    x_drivetrainList = x_drivetrainList & ">" & "FWD" & "</option>"
    x_drivetrainList = x_drivetrainList & "<OPTION value=""RWD"""
    If x_drivetrain = "RWD" Then
        x_drivetrainList = x_drivetrainList & " selected"
    End If
    x_drivetrainList = x_drivetrainList & ">" & "RWD" & "</option>"
    x_drivetrainList = x_drivetrainList & "<OPTION value=""AWD"""
    If x_drivetrain = "AWD" Then
        x_drivetrainList = x_drivetrainList & " selected"
    End If
    x_drivetrainList = x_drivetrainList & ">" & "AWD" & "</option>"
x_drivetrainList = x_drivetrainList & "</select>"
response.write x_drivetrainList
%>
</font>&nbsp;</td>
<td bgcolor="white" width="156" height="25">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156" height="25">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156">&nbsp;</td>
<td bgcolor="white" width="156">&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156">                <p><font face="Arial"><b><span style="font-size:10pt;">Features</span></b></font></p>
</td>
<td bgcolor="white" width="468" colspan="3"><p><font face="Arial"><span style="font-size:8pt;"><textarea cols="40" rows="7" name="x_features"><%= x_features %></textarea></span></font></td>
</tr>
<tr>
<td bgcolor="white" width="156"><b><span style="font-size:10pt;"><font face="Arial">&nbsp;</font></span></b></td>
<td bgcolor="white" width="156">&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="624" colspan="4"><img src="images/photosBG4.gif" width="600" height="22" border="0"></td>
</tr>
<tr>
<td bgcolor="white" width="156">                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_1) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=1" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=1" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_2) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=2" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=2" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_3) Then %>
<font face="Arial" size="2"><a href="car_photo_3_bv.asp?key=<%= key %>" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=3" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_4) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=4" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=4" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><input type="file" name="x_photo_1" onchange="if (this.form.a_x_photo_1[2]) this.form.a_x_photo_1[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_2" onchange="if (this.form.a_x_photo_2[2]) this.form.a_x_photo_2[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_3" onchange="if (this.form.a_x_photo_3[2]) this.form.a_x_photo_3[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_4" onchange="if (this.form.a_x_photo_4[2]) this.form.a_x_photo_4[2].checked=true;" size="5"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_1) Then %>
<input type="radio" name="a_x_photo_1" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_1" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_1" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_2) Then %>
<input type="radio" name="a_x_photo_2" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_2" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_2" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_3) Then %>
<input type="radio" name="a_x_photo_3" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_3" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_3" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_4) Then %>
<input type="radio" name="a_x_photo_4" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_4" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_4" value="3">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_5) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=5" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=5" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p align="center">
<% If not isnull(x_photo_6) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=6" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=6" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_7) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=7" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=7" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_8) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=8" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=8" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_5" onchange="if (this.form.a_x_photo_5[2]) this.form.a_x_photo_5[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_6" onchange="if (this.form.a_x_photo_6[2]) this.form.a_x_photo_6[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_7" onchange="if (this.form.a_x_photo_7[2]) this.form.a_x_photo_7[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_8" onchange="if (this.form.a_x_photo_8[2]) this.form.a_x_photo_8[2].checked=true;" size="5"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_5) Then %>
<input type="radio" name="a_x_photo_5" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_5" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_5" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_6) Then %>
<input type="radio" name="a_x_photo_6" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_6" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_6" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_7) Then %>
<input type="radio" name="a_x_photo_7" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_7" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_7" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_8) Then %>
<input type="radio" name="a_x_photo_8" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_8" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_8" value="3">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">
&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">
                            <p align="center"><% If not isnull(x_photo_9) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=9" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=9" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p align="center">
<% If not isnull(x_photo_10) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=10" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=10" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_11) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=11" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=11" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_12) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=12" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=12" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_9" onchange="if (this.form.a_x_photo_9[2]) this.form.a_x_photo_9[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><input type="file" name="x_photo_10" onchange="if (this.form.a_x_photo_10[2]) this.form.a_x_photo_10[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_11" onchange="if (this.form.a_x_photo_11[2]) this.form.a_x_photo_11[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_12" onchange="if (this.form.a_x_photo_12[2]) this.form.a_x_photo_12[2].checked=true;" size="5"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_9) Then %>
<input type="radio" name="a_x_photo_9" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_9" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_9" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_10) Then %>
<input type="radio" name="a_x_photo_10" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_10" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_10" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_11) Then %>
<input type="radio" name="a_x_photo_11" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_11" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_11" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_12) Then %>
<input type="radio" name="a_x_photo_12" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_12" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_12" value="3">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">
&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">
                            <p align="center"><% If not isnull(x_photo_13) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=13" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=13" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p align="center">
<% If not isnull(x_photo_14) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=14" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=14" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_15) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=15" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=15" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_16) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=16" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=16" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_13" onchange="if (this.form.a_x_photo_13[2]) this.form.a_x_photo_13[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><input type="file" name="x_photo_14" onchange="if (this.form.a_x_photo_14[2]) this.form.a_x_photo_14[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_15" onchange="if (this.form.a_x_photo_15[2]) this.form.a_x_photo_15[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_16" onchange="if (this.form.a_x_photo_16[2]) this.form.a_x_photo_16[2].checked=true;" size="5"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_13) Then %><input type="radio" name="a_x_photo_13" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_13" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_13" value="3">
<% End If %>
</font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_14) Then %>
<input type="radio" name="a_x_photo_14" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_14" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_14" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_15) Then %>
<input type="radio" name="a_x_photo_15" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_15" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_15" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_16) Then %>
<input type="radio" name="a_x_photo_16" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_16" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_16" value="3">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">
&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">
                            <p align="center"><% If not isnull(x_photo_17) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=17" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=17" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p align="center">
<% If not isnull(x_photo_18) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=18" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=18" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_19) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=19" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=19" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_20) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=20" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=20" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_17" onchange="if (this.form.a_x_photo_17[2]) this.form.a_x_photo_17[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_18" onchange="if (this.form.a_x_photo_18[2]) this.form.a_x_photo_18[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_19" onchange="if (this.form.a_x_photo_19[2]) this.form.a_x_photo_19[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_20" onchange="if (this.form.a_x_photo_20[2]) this.form.a_x_photo_20[2].checked=true;" size="5"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_17) Then %>
<input type="radio" name="a_x_photo_17" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_17" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_17" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_18) Then %>
<input type="radio" name="a_x_photo_18" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_18" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_18" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_19) Then %>
<input type="radio" name="a_x_photo_19" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_19" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_19" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_20) Then %>
<input type="radio" name="a_x_photo_20" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_20" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_20" value="3">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">
&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">
                            <p align="center"><% If not isnull(x_photo_21) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=21" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=21" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p align="center">
<% If not isnull(x_photo_22) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=22" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=22" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_23) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=23" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=23" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_24) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=24" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=24" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_21" onchange="if (this.form.a_x_photo_21[2]) this.form.a_x_photo_21[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><input type="file" name="x_photo_22" onchange="if (this.form.a_x_photo_22[2]) this.form.a_x_photo_22[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_23" onchange="if (this.form.a_x_photo_23[2]) this.form.a_x_photo_23[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_24" onchange="if (this.form.a_x_photo_24[2]) this.form.a_x_photo_24[2].checked=true;" size="5"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_21) Then %>
<input type="radio" name="a_x_photo_21" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_21" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_21" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_22) Then %>
<input type="radio" name="a_x_photo_22" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_22" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_22" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_23) Then %>
<input type="radio" name="a_x_photo_23" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_23" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_23" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_24) Then %>
<input type="radio" name="a_x_photo_24" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_24" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_24" value="3">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">
&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">
                            <p align="center"><% If not isnull(x_photo_25) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=25" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=25" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p align="center">
<% If not isnull(x_photo_26) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=26" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=26" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_27) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=27" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=27" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">                            <p align="center"><% If not isnull(x_photo_28) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=28" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=28" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_25" onchange="if (this.form.a_x_photo_25[2]) this.form.a_x_photo_25[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><input type="file" name="x_photo_26" onchange="if (this.form.a_x_photo_26[2]) this.form.a_x_photo_26[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_27" onchange="if (this.form.a_x_photo_27[2]) this.form.a_x_photo_27[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_28" onchange="if (this.form.a_x_photo_28[2]) this.form.a_x_photo_28[2].checked=true;" size="5"></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_25) Then %>
<input type="radio" name="a_x_photo_25" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_25" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_25" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_26) Then %>
<input type="radio" name="a_x_photo_26" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_26" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_26" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_27) Then %>
<input type="radio" name="a_x_photo_27" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_27" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_27" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_28) Then %>
<input type="radio" name="a_x_photo_28" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_28" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_28" value="3">
<% End If %></font></td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">
&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156" height="100">
                            <p align="center"><% If not isnull(x_photo_29) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=29" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=29" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p align="center">
<% If not isnull(x_photo_30) Then %>
<font face="Arial" size="2"><a href="car_ph.asp?key=<%= key %>&nr=30" target="_self"><img src="car_ph.asp?key=<%= key %>&nr=30" border=0 width="100"></a><%else %><img src="images/emptyCarPhoto.gif" width="100" height="75" border="0">
<% End If %></font></td>
<td bgcolor="white" width="156" height="100">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156" height="100">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><input type="file" name="x_photo_29" onchange="if (this.form.a_x_photo_29[2]) this.form.a_x_photo_29[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><input type="file" name="x_photo_30" onchange="if (this.form.a_x_photo_30[2]) this.form.a_x_photo_30[2].checked=true;" size="5"></font></td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center"><font face="Arial" size="2"><% If not isnull(x_photo_29) Then %><input type="radio" name="a_x_photo_29" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_29" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_29" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p align="center">
<font face="Arial" size="2"><% If not isnull(x_photo_30) Then %>
<input type="radio" name="a_x_photo_30" value="2">Remove&nbsp;<input type="radio" name="a_x_photo_30" value="3">Replace<br>
<% Else %>
<input type="hidden" name="a_x_photo_30" value="3">
<% End If %></font></td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="156">
                            <p align="center">&nbsp;</td>
<td bgcolor="white" width="156">
&nbsp;</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="156">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="624" colspan="4">

&nbsp;</td>
</tr>
</table>
<p align="center">
<input type="submit" name="Action" value="UPDATE ">
</form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="adminlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/back.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;&nbsp;</b></font><a href="adminlist.asp"><font face="Arial" size="2" color="black"><b>Back to Inventory List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
