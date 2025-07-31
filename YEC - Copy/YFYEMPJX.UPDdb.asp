<%@LANGUAGE="VBSCRIPT"  codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Expires = 0
Response.Buffer = true 

JXYM = REQUEST("JXYM")
g1 = REQUEST("g1")
s1  = REQUEST("s1")  
w1= REQUEST("w1")  

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

tmpRec = Session("empjxedit")  


Set conn = GetSQLServerConnection()	   
conn.BeginTrans 

				
for i = 1 to TotalPage 
	for j = 1 to PageRec  
		op = request("op")(j)
		descp = request("descp")(j)
		HXSL = request("HXSL")(j)
		HESO = request("HESO")(j)
		per = request("per")(j) 
		jxwhsno = request("jxwhsno")(j) 
		if op="del" then 
			sql="delete YFYMJIXO where jxym='"& jxym &"' and groupid='"& tmpRec(i, j, 3) &"' and shift='"& tmpRec(i, j, 4) &"' "&_
				"and stt='"& tmpRec(i, j, 5) &"' and autoid='"& tmpRec(i, j, 1) &"' "  
			'response.write sql &"<BR>"	
			conn.execute(Sql)
		else		
			sql="update YFYMJIXO set jxwhsno='"& jxwhsno &"' , descp=N'"& descp  &"' ,   "&_
				"HXSL='"& HXSL  &"' , HESO='"& HESO  &"' , "&_
				"per='"& per &"' "&_
				"where jxym='"& jxym &"' and groupid='"& tmpRec(i, j, 3) &"' and shift='"& tmpRec(i, j, 4) &"' "&_
				"and zuno='"& tmpRec(i, j, 11) &"' and stt='"& tmpRec(i, j, 5) &"' and autoid='"& tmpRec(i, j, 1) &"' "  
			response.write sql &"<BR>"	
			conn.execute(sql)
		end if 	
	next
next 	
 
'response.end 

if conn.Errors.Count = 0 then 
	conn.CommitTrans
	Set conn = Nothing		
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YFYEMPJX.foreedit.asp?jxym="&"<%=jxym%>"&"&G1="&"<%=g1%>"&"&S1="&"<%=s1%>"&"&w1="&"<%=w1%>" , "_self" 
	</script>	
<%  
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YFYEMPJX.foreedit.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 