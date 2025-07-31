<%@LANGUAGE="VBSCRIPT" codepage=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""
session.codepage=950 
'Response.Expires = 0
'Response.Buffer = true

Set CONN = GetSQLServerConnection()
conn.BeginTrans  
for i = 1 to 10  
	empid = trim(request("empid")(i))
	country = trim(request("country")(i))
	visano = trim(request("visano")(i))
	sdat = trim(request("dat1")(i))
	edat = trim(request("dat2")(i))	
	visaamt = trim(request("visaamt")(i))
	memo = trim(request("memo")(i))
	
	if empid<>"" and  sdat<>"" and edat<>"" then  
		sql="insert into empVisadata ( country, empid, visano, sdat, edat,  visaAmt, memo,  mdtm, muser ) values ( "&_
			"'"& country &"', '"& empid &"' , '"& visaNo &"', '"& sdat &"' , '"& edat &"' , '"& visaAmt &"', N'"& memo &"' , "&_
			"getdate(), '"& session("NetUser") &"' ) "
		response.write sql&"<BR>"	
		conn.execute(Sql)
	end if 		
	
next 


if conn.Errors.Count = 0 then 
	conn.CommitTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功DATA CommitTrans SUCCESS!!"
		OPEN "empbe03.asp" , "_self" 
	</script>	
<%
	
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗DATA CommitTrans ERROR !!"
		OPEN "empbe03.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%> 
 
