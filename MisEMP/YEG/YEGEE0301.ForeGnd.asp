<%@ Language=VBScript codepage=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%

self="YEGEE0301"
dat1 = (trim(request("dat1")))
dat2 = (trim(request("dat2")) )
eid  = request("eid")


response.write dat1 & "  ----" & dat2 
'response.end 

D1 = replace(dat1, "/", "" )
D2 = replace(dat2, "/", "" )   
df = datediff("d",dat1,dat2) 
eid= Ucase(trim(request("eid")))
 
aidstr = session("netuser") & minute(now())&second(now()) 
    
Set conn = GetSQLServerConnection() 
response.write conn  &"<BR>" 
'conn.BeginTrans 

'---step1  
sql_1="exec A_tranEmpData  '"& eid &"' , '"& dat1 &"', '"&  dat2  &"' " 
response.write sql_1
conn.execute(Sql_1)  
'response.write sql_1 &"<BR>"
 
'response.end  
'--step2 
sql2="exec A_transRptDay '"& dat1 &"', '"&  dat2  &"' ,'"& eid &"'" 
conn.execute sql2  
response.write sql2 &"<BR>"
 
'response.end 
'response.write  err.number


'if  conn.Errors.Count=0 or err.number=0  then 
	'conn.CommitTrans 
if err.number=0  then 
	'conn.CommitTrans  
%>	<script language=vbscript>
		alert "Data CommitTrans sunccess!!  OK!!"
		open "<%=self%>.asp" , "_parent"
	</script>
<%	response.end 
else
	'conn.RollbackTrans	 
%>
	<script language=vbscript>
		alert "Data CommitTrans Fail!!  Error!!"
		open "<%=self%>.asp" , "_parent"
	</script>	
<%  response.end 
end if 	
%>