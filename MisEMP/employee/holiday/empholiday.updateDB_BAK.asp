<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
<%
Response.Expires = 0
Response.Buffer = true 

empid=reuqets("empid")
HOLIDAY_TYPE  = request("HOLIDAY_TYPE")
HHDAT1 = request("HHDAT1")
HHDAT2 = request("HHDAT2")
HHTIM1 = request("HHTIM1")
HHTIM2 = request("HHTIM2")
toth = request("toth")
memo = request("memo")


Set CONN = GetSQLServerConnection()  
conn.BeginTrans
 
sql="insert into empHoliday ( empid, jiaType, DateUP, TimeUP, DateDown, TimeDowm, HHour, memo, Muser ) values ( "&_
	"'"& empid &"', '"& HOLIDAY_TYPE &"', '"& HHDAT1 &"', '"& HHTIM1 &"', '"& HHDAT2 &"', '"& HHTIM2 &"', "&_
	"'"& toth &"', '"& memo &"', '"& session("NETUSER") &"' ) 
response.write sql 	
RESPONSE.END 

if conn.Errors.Count = 0 then 
	conn.CommitTrans	
	'response.redirect "empfile.salary.asp?empidstr=" & empidstr 
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans  SUCCESS!!"		
	</script>		
<%
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		OPEN "empfile.salary.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 