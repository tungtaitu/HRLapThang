<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true

Set CONN = GetSQLServerConnection()
conn.BeginTrans  
for i = 1 to 10  
	empid = trim(request("empid")(i))
	country = trim(request("country")(i))
	indate = trim(request("indate")(i))
	sdat = trim(request("dat1")(i))
	edat = trim(request("dat2")(i))	
	memo = trim(request("memo")(i))
	
	if empid<>"" and  sdat<>"" and edat<>"" then  
		sql="insert into empHTdata ( country, empid, indat, sdat, edat, memo,  mdtm, muser ) values ( "&_
			"'"& country &"', '"& empid &"' , '"& indate &"', '"& sdat &"' , '"& edat &"' ,N'"& memo &"' , "&_
			"getdate(), '"& session("NetUser") &"' ) "
		response.write sql&"<BR>"	
		conn.execute(Sql)
	end if 		
	
next 


if conn.Errors.Count = 0 then 
	conn.CommitTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功DATA CommitTrans SUCCESS!!"
		OPEN "empbe04.asp" , "_self" 
	</script>	
<%
	
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗DATA CommitTrans ERROR !!"
		OPEN "empbe04.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%> 
 
