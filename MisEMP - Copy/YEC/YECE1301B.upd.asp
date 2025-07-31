<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YECE1301B"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()


years = request("years")
Ct = request("ct")
whsno = request("whsno") 

' response.write "years=" & firstday &"<BR>"
' response.write "endday=" & endday &"<BR>" 
' response.write "cfg=" & cfg 
'response.end 
Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
 
nam=years
country=ct
whsno=whsno 
	
for i = 1 to pagerec	
	if trim(ct)="" then country=Ucase(trim(request("country")(i))) 	
	grade=Ucase(trim(request("grade")(i)))
	days=Ucase(trim(request("days")(i)))
	hs=Ucase(trim(request("hs")(i)))
	memos=Ucase(trim(request("memos")(i)))
	aid=Ucase(trim(request("aid")(i)))
	kj=Ucase(trim(request("kj")(i)))
	whsno=Ucase(trim(request("w1")(i)))

	if  grade<>"" and days<>"" and nam<>"" and country<>""  then 
		sql2="select* from empnzjj_set where years='"& nam &"' and whsno='"& whsno &"' and country='"& country &"' and grade='"& grade &"' " 
		Set rs = Server.CreateObject("ADODB.Recordset") 
		rs.open sql2, conn, 3, 3 
		response.write sql2&"<BR>"
		if rs.eof  then 
			sql="insert into [EMPNZJJ_set]([whsno], [years], [country], [grade], [days], [hs], [memos], "&_
					"[keyinDate], [keyinBy], [mdtm], [muser] ,kj ) values( "&_
					"'"&whsno&"','"&nam&"','"&country&"','"&grade&"','"&days&"','"&hs&"','"&memos&"', "&_
					"getdate(),'"&session("netuser")&"',getdate(),'"&session("netuser")&"','"&kj&"' ) "
		else
			sql="update [EMPNZJJ_set] set  whsno='"&whsno &"'  ,"&_
					"grade = '"&grade&"', days='"&days&"', hs='"&hs&"',  memos=N'"&memos&"',kj='"&kj&"', "&_
					"mdtm=getdate(),muser='"&session("netuser")&"' "&_
					"where aid='"& aid &"'  "
		end if 			
		set rs=nothing 
		conn.execute(sql) 
		response.write sql&"<BR>" 	
	end if 
 
next
response.write err.number &"<BR>"
response.write conn.errors.count &"<BR>"

for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next  

'RESPONSE.END

if err.number = 0 then
	conn.CommitTrans
	Set Session("YECE03") = Nothing
	Set conn = Nothing 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.fore.asp" , "_self"
	</script>
<% 
ELSE
	conn.RollbackTrans 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.fore.asp" , "_self"
	</script>
<%  response.end
END IF
%>
 