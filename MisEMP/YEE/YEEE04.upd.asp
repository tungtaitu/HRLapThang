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
sts2 = request("sts2") 

dd1=request("dd1")
dd2=request("dd2")
sothe=ucase(triM(request("sothe")))
ct = request("ct")
otherQ = trim(request("otherQ")) 


'response.write DAT1 & DAT2 & whsno & groupid & country 
tmpRec = Session("YEBE0104B")   
 
pagerec=request("pagerec") 
totalpage=request("totalpage")
zjdays = request("zjdays") 
'response.write cdbl(pagerec)+cdbl(zjdays)
'response.end 
conn.BeginTrans 
x = 0 
f_xid= ""
xidstr=""
for i = 1   to totalpage 
	for j = 1 to  pagerec  
		'response.write tmpRec(i, j, 0) &"<BR>" 		
		op = tmpRec(i, j, 0)  
		xid = tmpRec(i, j, 1)   
		empid = tmpRec(i, j, 2)
		s_dat = tmpRec(i, j, 6)
		e_dat = tmpRec(i, j, 8)
			
		if op="Y" then 
			sql="update empholiday set xjsts='*' where isnull(xid,'')='"& xid &"' and empid='"& empid &"'  and convert(char(10), dateup,111) between '"& s_dat &"' and '"& e_dat &"' " 
			response.write sql &"<BR>" 
			conn.execute(sql)
		end if 	
	next 	
next  

 
' response.write y &"<BR>"
'response.end   
' for zz = 0 to ubound(split(f_xid,","))	
	' n_xid = split(f_xid,",")(zz)	
	' if trim(n_xid)<>""     then 
		 ' sqlx="select jiatype, xid , min(convert(char(10),dateup,111) as mindat  , max(convert(char(10),dateup,111) as maxdat "&_
				  ' "from empHoliday where isnull(xid,'')='"& n_xid  &"' and group by jiatype, xid "
		 ' response.write " f .sql="& sqlx &"<BR>"
	' end if 
' next 	

'response.end 
'response.redirect "empbasicB.fore.asp?yymm="& yymm 



if ( conn.Errors.Count = 0 or err.number=0 )  then 
	conn.CommitTrans	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>		
		open "yeee04.Fore.asp?sts2="&sts2&"&dd1="& "<%=dd1%>" &"&dd2="& "<%=dd2%>" &"&sothe="& "<%=sothe%>" &"&ct="&  "<%=ct%>" &"&otherQ="& "<%=otherQ%>"  , "_self" 
	</script>	
<%ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		open "yeee04.Fore.asp?dd1="& "<%=dd1%>" &"&dd2="& "<%=dd2%>" &"&sothe="& "<%=sothe%>" &"&ct="&  "<%=ct%>" &"&otherQ="& "<%=otherQ%>"  , "_self" 
	</script>	
<%	response.end 
END IF  
%>
  
 