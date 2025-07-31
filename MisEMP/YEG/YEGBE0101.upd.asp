<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file="../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true  

self="YEGBE0101"

Set CONN = GetSQLServerConnection()   
Days= request("workdays")
yymm= request("yymm")
groupid = request("groupid")
shift = request("shift")
zuno = request("zuno")

'response.write groupid &"-"&zuno&"-"&shift
For i = 1 to Days    
	if groupid="A061" then 
		calcH=8
	elseif request("uptim")(i)>="17:00" then 
		calcH=8.5
	else
		calcH=9
	end if  
	
	uptim = request("uptim")(i)	
	dat = request("dat")(i)	
	if dat<>"" then 	
		sqla="delete empDS where convert(char(10), dat,111)='"& request("dat")(i) &"' "&_
			 "and groupid='"& groupid &"' and zuno='"& zuno &"' and shift='"& shift &"' "	
		conn.execute(sqla) 	
			
		sqlb="insert into empDS (groupid, zuno, shift, dat, Uptim, calcH,mdtm, muser , userIP ) values ( "&_
			 "'"& groupid &"', '"& zuno &"', '"& shift &"', '"& dat &"', '"& uptim  &"', '"& calcH &"',  "&_
			 "getdate(), '"& session("NETUSER") &"', '"& session("vnlogIP") &"' ) " 
		conn.execute(sqlb)	
	end if	
Next 
response.write "..."
response.redirect "YEGBE0101.fore.asp?yymm="& yymm &"&groupid=" & groupid &"&zuno=" & zuno &"&shift=" & shift 
'response.end 

%>
 