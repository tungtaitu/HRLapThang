<%@LANGUAGE="VBSCRIPT" codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../../GetSQLServerConnection.fun" --> 

<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true 

Set CONN = GetSQLServerConnection()   
 
DAT1 = REQUEST("DAT1")
DAT2 = REQUEST("DAT2")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
country = trim(request("country"))  
QUERYX = trim(request("empid1"))  

'response.write DAT1 & DAT2 & whsno & groupid & country 
tmpRec = Session("yeee0401B")   
 
pagerec=request("pagerec") 
gTotalPage=request("gTotalPage")
zjdays = request("zjdays") 
'response.write cdbl(pagerec)+cdbl(zjdays)
'response.end 
conn.BeginTrans 
x = 0 
f_xid= ""
xidstr="" 

for i = 1   to gTotalPage 
	for j = 1 to  pagerec  
		'response.write tmpRec(i, j, 0) &"<BR>" 		
		op = tmpRec(i, j, 0)  
		years = trim(tmpRec(i, j, 1))
		country = tmpRec(i, j,2)
		phi_vnd=replace(trim(tmpRec(i, j, 3) ),",","")
		phi_usd=replace(trim(tmpRec(i, j, 4) ),",","")
		btax =replace(trim(tmpRec(i, j, 6)),",","")
		if years<>"" and country<>"" and phi_vnd<>"" then 
			sqlx="select * from  empphi where years='"&years&"' and country='"&country&"' "
			Set rs = Server.CreateObject("ADODB.Recordset")  
			rs.open sqlx , conn , 1, 3 
			if rs.eof then 
				sql="insert into  empphi ( years, country, phi_vnd, phi_usd, mdtm, muser, btax   ) values ( "&_
						"'"&years&"','"&country&"','"&phi_vnd&"','"&phi_usd&"' , getdate(), '"&session("netuser") &"','"& btax &"'  ) "
				conn.execute(sql)	
				response.write sql&"<br>" 
			else
				sql="update empphi set btax='"& btax &"', phi_vnd='"&phi_vnd&"' , phi_usd='"&phi_usd&"' , mdtm=getdate() , muser='"&session("netuser") &"' where years='"&years&"' and country='"&country&"' "
				conn.execute(Sql) 
				response.write sql&"<br>" 
			end if 	
			set rs=nothing 
		end if 			
	next 	
next  
 
'response.end 
'response.redirect "empbasicB.fore.asp?yymm="& yymm 



if ( conn.Errors.Count = 0 or err.number=0 )  then 
	conn.CommitTrans	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>		
		open "yeee0401.fore.asp" , "_self" 
	</script>	
<%ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		open "yeee0401.fore.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
  
 