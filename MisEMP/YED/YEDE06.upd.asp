<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YEDE06"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

 
 
Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR="" 
for i = 1 to pagerec 
	yymm=request.form("yymm")(i)
	groupid=request.form("groupid")(i)
	zuno=request.form("zuno")(i)
	shift=request.form("shift")(i) 
	tothrs=request.form("tothr")(i) 
	if tothrs="" then tothr = "0"    	
	sqla="delete emptothr where yymm='"&yymm&"'  and groupid='"&groupid&"' and zuno='"&zuno&"' and shift='"&shift&"' " 
	'response.write sqla
	conn.execute(sqla) 
	sqlb="insert into emptothr (yymm,groupid,zuno,shift,tothrs,mdtm,muser) values ('"&yymm&"','"&groupid&"','"&zuno&"','"&shift&"','"&tothrs&"',getdate(),'"&session("netuser")&"'  )" 	
	'response.write sqlb
	'response.end
	conn.execute(sqlb)	
	x = x+1 
next  

'response.end  
'response.write err.number &"<BR>"
'response.write conn.errors.count &"<BR>"
for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next   
'RESPONSE.END

if err.number = 0   then
	conn.CommitTrans
	 
	Set conn = Nothing 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.fore.asp?d1="&"<%=yymm%>" , "Fore"
	</script>
<% 
ELSE
	conn.RollbackTrans 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "empsalaryHW.asp" , "_self"
	</script>
<%  response.end
END IF
%>
 