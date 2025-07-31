<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!----- #include file="../ADOINC.inc" ------>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<html>

<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<body>
<%
Response.Buffer = true
Response.Expires = 0
session.codepage="65001"
%>

<%
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")
yymm = request("yymm") 
'response.write gTotalPage &"<BR>"
NOWMONTH=CSTR(YEAR(DATE()))&RIGHT("00"&CSTR(MONTH(DATE())),2)

Set conn = GetSQLServerConnection()
tmpRec = Session("YEBBB0101")

conn.BeginTrans
x=0
y=""
for i = 1 to TotalPage 
	for j = 1 to PageRec 
		empinfo=trim(tmpRec(i, j, 2))&" "&trim(tmpRec(i, j, 3)) 
		if tmpRec(i, j , 0) = "del" then		
			sql="delete  basicCode where  autoid='"& trim(tmpRec(i, j, 4)) &"' "	
			conn.execute(sql)	
		else
			if trim(tmpRec(i, j, 0))="upd" or yymm<>NOWMONTH then 
				empid=trim(tmpRec(i, j, 1))
				country=trim(tmpRec(i, j, 4))
				whsno=trim(tmpRec(i, j, 7))
				job=trim(tmpRec(i, j, 6))
				memo=trim(tmpRec(i, j, 23))
				sql="select * from BempJ where yymm='"& yymm &"' and empid='"& empid &"' and whsno='"& whsno &"'  "
				Set rds = Server.CreateObject("ADODB.Recordset")
				rds.open sql, conn, 3, 3 
				if rds.eof then 
					sqla="insert into BempJ ( yymm, empid, whsno, country, job, memo, mdtm, muser ) values ( "&_
						 " '"& yymm &"','"& empid  &"','"& whsno &"','"& country  &"','"& job &"', '"& memo &"', "&_
						 "getdate() ,'"& Session("NETUSER") &"' ) " 
				else
					sqla="update  BempJ set job='"& job &"' , memo ='"& memo &"' , mdtm=getdate(), muser='"& session("netUser") &"' "&_
						 "where yymm='"& yymm &"' and empid='"& empid&"' and whsno='"& whsno &"'   "
				end if 
				'response.write sqla &"<BR>"
				conn.execute(sqla)
		

				sqlT="SELECT * FROM BASICCODE WHERE  FUNC='LEV' and sys_type='"& job &"' "
				set rsTmp=conn.execute(SqlT)
				tmpRec(i, j, 25) = rsTmp("sys_value")
				x=x+1
				y=y& X &". "&trim(tmpRec(i, j, 13))&" "&empid&" "&empinfo& "<BR>異動日期: " & yymm &"<br>異動內容: " & trim(tmpRec(i, j, 24))&trim(tmpRec(i, j, 15))&" → " & trim(tmpRec(i, j, 6))&tmpRec(i, j, 25) &"<BR>異動說明:<BR>" & trim(tmpRec(i, j, 23)) &"<hr size=0	style='border: 1px dotted #999999;' align=left >"
			end if	
		end if 		
	next
next 
'response.write Y 
'response.end 
if x=0 then 
	response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td align=center>"
	response.write "<font class=txt12 color=red>無修改資料!!</font><BR><BR>"	
	response.write "</td></tr>"
	response.write "<tr><td align=center><BR><BR><a href='YEBBB0101.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
	response.write "</td></tr></table>"
	response.end 
else
	if  conn.errors.count=0  then 
		conn.CommitTrans	
		conn.close
		set	conn=nothing
		'response.redirect "yeaae0101.Fore.asp?A1="&A1
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"
		response.write "<font class=txt12 color=red>資料處理成功OK(SUCCESS)!!</font><BR><BR>"
		response.write y
		response.write "</td></tr>"
		response.write "<tr><td align=center><BR><BR><a href='YEBBB0101.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"
		set Session("YEBBB0101") = nothing 
		response.end 
		
	else	
		conn.close
		set	conn=nothing
		response.write "errors!!"
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"	
		response.write "</td></tr>"
		response.write "<tr><td><BR><BR><a href='YEBBB0101.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"	
		Response.End 
	end if 	 
end if 	
%>
</body>
</html>