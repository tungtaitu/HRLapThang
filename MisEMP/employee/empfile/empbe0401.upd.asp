<%@LANGUAGE="VBSCRIPT" codepage=950 %>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""
session.codepage=950
'Response.Expires = 0
'Response.Buffer = true 

RecordInDB = request("RecordInDB")

s_empid=request("s_empid")
s_dat1=request("s_dat1")
s_dat2=request("s_dat2")
s_country=request("s_country")

'response.write s_country 
'response.end 
  

Set CONN = GetSQLServerConnection()
conn.BeginTrans  
for i = 1 to RecordInDB+1    
	aid = request("aid")(i)	
	sdat = trim(request("dat1")(i))
	edat = trim(request("dat2")(i))		
	memo = trim(request("memo")(i)) 
	
	sql="update EmpHTdata set  sdat='"& sdat &"', edat='"& edat &"', memo='"& memo&"'  "&_
		"where  aid='"& aid &"' "
	conn.execute(Sql)	
	
next 


if conn.Errors.Count = 0 then 
	conn.CommitTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功DATA CommitTrans SUCCESS!!"
		OPEN "empbe0401.Fore.asp?s_empid="& "<%=s_empid%>" &"&s_dat1=" & "<%=s_dat1%>" &"&_sdat2=" & "<%=s_dat2%>" &"&s_country=" & "<%=s_country%>"  , "_self"  
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
 
