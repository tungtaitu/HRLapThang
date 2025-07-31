<%@LANGUAGE="VBSCRIPT" codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""
'session.codepage=65001
'Response.Expires = 0
'Response.Buffer = true 

RecordInDB = request("RecordInDB")

s_empid=request("s_empid")
s_dat1=request("s_dat1")
s_dat2=request("s_dat2")
s_country=request("s_country")

'response.write s_country 
'response.end 
  
xx = 0 
Set CONN = GetSQLServerConnection()
conn.BeginTrans  
for i = 1 to RecordInDB+1    
	aid = request("aid")(i)
	visano = trim(request("visano")(i))
	sdat = trim(request("dat1")(i))
	edat = trim(request("dat2")(i))	
	visaamt = trim(request("visaamt")(i))
	memo = trim(request("memo")(i)) 
	op =  trim(request("op")(i)) 
	
	if op="upd" then 
		sql="update empVisaData set visano='"& visano &"', sdat='"& sdat &"', edat='"& edat &"', visaamt='"& visaAmt &"', memo=N'"& memo&"'  "&_
			"where  aid='"& aid &"' "
		conn.execute(Sql)	
		response.write sql 
		xx = xx+1
	end if 	
next 
if xx = 0 then %>
	<SCRIPT LANGUAGE=VBSCRIPT>	
		OPEN "empbe0301.Fore.asp?s_empid="& "<%=s_empid%>" &"&s_dat1=" & "<%=s_dat1%>" &"&_sdat2=" & "<%=s_dat2%>" &"&s_country=" & "<%=s_country%>"  , "_self"  
	</script>	
<%end if 
'response.end 
if conn.Errors.Count = 0 then 
	conn.CommitTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功DATA CommitTrans SUCCESS!!"
		OPEN "empbe0301.Fore.asp?s_empid="& "<%=s_empid%>" &"&s_dat1=" & "<%=s_dat1%>" &"&_sdat2=" & "<%=s_dat2%>" &"&s_country=" & "<%=s_country%>"  , "_self"  
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
 
