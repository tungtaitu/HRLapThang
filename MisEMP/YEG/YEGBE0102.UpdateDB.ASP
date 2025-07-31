<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">

<%
Response.Buffer = true
Response.Expires = 0
%>

<%
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

self="YEGBE0102"
Set conn = GetSQLServerConnection()
tmpRec = Session("YEGBE0102")
'on error resume next
'conn.BeginTrans
x = 0
y = ""
errmsg=""
for i = 1 to gTotalPage 
	for j = 1 to PageRec 
		if trim(tmpRec(i, j, 7))<>"" then 	
			if trim(tmpRec(i, j, 0))= "upd" then 
				sql="insert into empscg (groupid, zuno, shift, func , mdtm, muser ) values ( "&_
					"'"& tmpRec(i, j, 1) &"','"& tmpRec(i, j, 2) &"','"& tmpRec(i, j, 3) &"','"& tmpRec(i, j, 7) &"', "&_
					"getdate(), '"& session("netuser") &"' ) " 			 
				conn.execute(sql)
				response.write sql &"<BR>"			
			end if	
		end if  			
	next
next 
 
if errmsg=""  then 
	'conn.CommitTrans
	Set Session("YEGBE0102") = Nothing 
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans Success (OK)!!"
		OPEN "<%=self%>.asp" , "_self" 
	</script>	
<%	
	
ELSE
	'conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
   


%>
