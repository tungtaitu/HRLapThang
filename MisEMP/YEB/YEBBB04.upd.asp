<%@Language=VBScript codepage=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
self="yebbb04"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")
whsno = request("whsno")
if whsno="" then whsno="LA"

Set conn = GetSQLServerConnection()
tmpRec = Session("yebbb04B")

'on error resume next
conn.BeginTrans
x = 0
 
for i = 1 to TotalPage 
	for j = 1 to PageRec  
		if tmpRec(i,j,0) = "del" then 
			sql="update emplicense set status='D', mdtm=getdate(), muser='"& session("netuser") &"'  where autoid='"& tmpRec(i,j,24) &"' "
			x=x+1
			response.write sql&"<BR>"		
			conn.execute(sql)
		else
			if tmpRec(i,j,0)="upd" and (trim(tmpRec(i,j,1))<>"" and tmpRec(i,j,2)<>"")  then 
				sql="insert into emplicense(cdwhsno, empid, licenseName, licenseOrg, licenseno, period_dat, carddata, "&_
					"qty, due_date, amt, cardmemo, mdtm, muser ) values ( "&_
					"'"&whsno&"','"&trim(tmpRec(i,j,1))&"','"&trim(tmpRec(i,j,2))&"','"&trim(tmpRec(i,j,3))&"',N'"&trim(tmpRec(i,j,18))&"', "&_
					"'"&trim(tmpRec(i,j,20))&"','"&trim(tmpRec(i,j,5))&"','"&trim(tmpRec(i,j,4))&"', "&_
					"'"&trim(tmpRec(i,j,6))&"','"&trim(tmpRec(i,j,21))&"','"&trim(tmpRec(i,j,22))&"', getdate(), "&_
					"'"& session("netuser") &"' ) " 
				response.write sql&"<BR>"			
				conn.execute(sql)
				x=x+1	
			end if 			
		end if  
	next
next 
'response.end 
 if err.number = 0 and x>"0"  then 
	conn.CommitTrans	
	Set Session("yebbb04B") = Nothing    
	Set conn = Nothing%>
	<script language=vbs>
		alert "資料處理成功 OK (data complete success)  "& <%=x%> &"  筆"
		open "<%=self%>.Fore.asp?whsno="&"<%=whsno%>", "_self"
	</script>
<%	
 else
	conn.RollbackTrans	
    Set Session("yebbb04B") = Nothing
	Set conn = Nothing
%>	<script language=vbs>
		alert "資料處理失敗(error) "& <%=x%> &" 筆"
	</script>
<% 	response.end 
 end if %>
