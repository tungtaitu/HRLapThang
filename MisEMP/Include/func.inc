<%
dd1=year(date())&"/"&right("00"&month(date()),2)&"/01"
dd2=year(date())&"/"&right("00"&month(date()),2)&"/"&right("00"&day(date()),2)

function formatdat(d) 
	response.write year(d)&"/"&right("00"&month(d),2)&"/"&right("00"&day(d),2) 
end function  
%>

<script language=vbscript> 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
	if window.event.keyCode = 113  then 
		go()
	end if 		
end function   
 

'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/"  & Right("00" & Month(d), 2) & "/" & Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________ 
</script> 