<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
<%
Response.Expires = 0
Response.Buffer = true 

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("empfilesalary") 

Set CONN = GetSQLServerConnection()  
conn.BeginTrans
x = 0
y = "" 
EMPIDSTR="" 
for i = 1 to TotalPage 
	for j = 1 to PageRec  
		'RESPONSE.WRITE TotalPage &"<br>"
		'RESPONSE.WRITE PageRec &"<br>"
		if trim(tmpRec(i, j, 1))<>"" then 
			IF trim(tmpRec(i, j, 0))="UPD" THEN 
				SQL="UPDATE EMPFILE SET BB='"& tmpRec(i, j, 19) &"', CV='"& tmpRec(i, j, 22) &"',"&_
					"PHU='"& tmpRec(i, j, 23) &"', NN='"& tmpRec(i, j, 24) &"', "&_
					"KT='"& tmpRec(i, j, 25) &"', MT='"& tmpRec(i, j, 26) &"' , "&_
					"TTKH='"& tmpRec(i, j, 27) &"',JOB='"& tmpRec(i, j, 6) &"', "&_
					"MDTM=GETDATE(), MUSER='"& SESSION("NETUSER") &"' "&_
					"WHERE EMPID='"& TRIM(tmpRec(i, j, 1)) &"' "  
				'RESPONSE.WRITE SQL &"<br>"	 
				EMPIDSTR = EMPIDSTR & "'" & TRIM(tmpRec(i, j, 1)) &"'," 
				'RESPONSE.WRITE EMPIDSTR &"<BR>"
				conn.execute(Sql)
			END  IF 
		END IF 
	next
next 	

'RESPONSE.END 

if conn.Errors.Count = 0 then 
	conn.CommitTrans
	Set Session("empfilesalary") = Nothing 	    
	response.redirect "empfile.salary.asp?empidstr=" & empidstr 
	Set conn = Nothing
	
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		OPEN "empfile.salary.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 