<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YECE1101"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()

cfg  = request("cfg") 

firstday  = request("calcdat")
endday = request("ccdt") 

' response.write "firstday=" & firstday &"<BR>"
' response.write "endday=" & endday &"<BR>" 
' response.write "cfg=" & cfg 
'response.end 
Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""

whsno = request("whsno")

for i = 1 to pagerec	
	empid=Ucase(trim(request("empid")(i)))
	person_qty=replace((trim(request("person_qty")(i))),",","")
	ut_mtax=replace((trim(request("ut_mtax")(i))),",","")
	tot_Mtax=replace((trim(request("tot_Mtax")(i))),",","")
	op=replace((trim(request("op")(i))),",","")
	serno=replace((trim(request("serno")(i))),",","")
	if (trim(request("person_qty")(i)))="" then person_qty=0 
	if (trim(request("ut_mtax")(i)))="" then ut_mtax=0 
	if (trim(request("tot_Mtax")(i)))="" then tot_Mtax=0 
	if empid<>""  then 
		sql="select * from empnotax where empid='"& empid &"' "
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql, conn, 1, 3 
		if rs.eof then 			
			sql="insert into empnotax(whsno,empid,person_qty ,ut_mtax ,tot_Mtax, mdtm, muser, keyindate, keyinby ) values ( "&_
					"'"&whsno&"','"&empid&"','"&person_qty&"','"&ut_mtax&"','"&tot_Mtax&"',getdate(),'"&session("netuser")&"',getdate(), '"&session("netuser")&"' ) "
			conn.execute(sql)
			'response.write sql &"<BR>"
			x = x+1			
		else	
			if rs("sts")="D" then op="E" 
			
			if cdbl(tot_Mtax)=0 then 
				'if op="D" then 
					sql="update empnotax set sts='D' , mdtm=getdate(), muser='"& session("netuser") &"' where empid='"& empid &"' "
					conn.execute(sql)
					'response.write sql &"<BR>"	
					x = x+1			
				'end if
			else	
				if op<>""   then 
					sql="update empnotax set person_qty='"& person_qty &"' , ut_mtax='"&ut_mtax&"' , tot_Mtax='"& tot_Mtax &"' , "&_
							"mdtm=getdate(), muser='"& session("netuser") &"' ,sts='' where empid='"& empid &"' "
					conn.execute(sql)
					'response.write sql &"<BR>"	
					x = x+1
				end if				
			end if 	
		end if 	
		rs.close 
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
		OPEN "<%=self%>.asp" , "_self"
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
 