<%
'*******************************************************************************************

Function Convert(ConvertString)

'*------------------------*
'* Convert "'"'s to "~"'s *
'*------------------------*
IF NOT ISNULL(ConvertString) THEN

	Convert=REPLACE(ConvertString,"'","''")

ELSE

	Convert="[Nothing]"	
	
END IF

End Function

'*******************************************************************************************

function ValueConvert(ConvertString)

'*------------------------*
'* Convert "'"'s to "~"'s *
'*------------------------*
IF NOT ISNULL(ConvertString) THEN

	IF ConvertString <> "" THEN

		ValueConvert=REPLACE(ConvertString,"'","&#039")

	END IF

END IF

end function

'*******************************************************************************************

function Dateconverter(datestring)


	dateconverter= day(datestring) & "/" & month(datestring) & "/" & year(datestring)
	
end function

'*******************************************************************************************

%>
