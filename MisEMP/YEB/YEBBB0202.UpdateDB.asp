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
NowMonth = request("NowMonth")

Set conn = GetSQLServerConnection()
tmpRec = Session("YEBBB0202")

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
			if trim(request("op")(j))="upd" then 
				empid=trim(tmpRec(i, j, 1))
				aid=tmpRec(i, j, 18)
				yymm=tmpRec(i, j, 25)				
				empid=trim(tmpRec(i, j, 1))
				country=trim(tmpRec(i, j, 4))
				whsno=trim(request("whsno")(j))				
				groupid=trim(request("groupid")(j))				
				zuno=trim(request("zuno")(j))								
				shift=trim(request("shift")(j))
				job=trim(tmpRec(i, j, 6))
				'Cdat= trim(tmpRec(i, j, 27)) 
				 
				memo=REPLACE(trim(request("memo")(j)),vbCrLf,"<BR>") 
				
				
				sqlt="select * from BempG where empid='"& empid &"' and  yymm = '"& yymm &"' "
				Set rst = Server.CreateObject("ADODB.Recordset")
				rst.open sqlt, conn, 1,3
				if rst.eof then 
					sqlz="insert into BempG (whsno,empid,country,groupid,zuno,shift, memo,mdtm,muser,yymm ) values ( "&_
						 "'"& whsno &"','"& empid &"','"& country &"', '"& groupid &"','"& zuno &"', '"& shift &"', "&_
						 "'"& memo &"',getdate() ,'"& session("NETuser") &"', '"& yymm &"' ) "					
					conn.execute(sqlz)	 				
				else 
					sqlz="update BempG set  whsno='"& whsno &"', groupid='"& groupid &"', zuno='"& zuno &"', shift='"& shift &"', "&_
						 "memo='"& memo &"',mdtm=getdate() , muser='"& session("NETuser") &"' "&_
						 "where empid='"& empid &"' and  yymm = '"& yymm &"' "	
					conn.execute(sqlz)	 
				end if	
				set rst=nothing 
				'response.write sqlz&"<BR>"
				
				if NowMonth=yymm then 
					'sqlxx="update empfile set whsno='"& whsno &"', groupid='"& groupid &"', zuno='"& zuno &"' , shift='"& shift &"', "&_
					'	  "mdtm=getdate(), muser='"& session("netuser") &"' where empid='"& empid &"' " 
					'response.write sqlxx &"<BR>"	  
					'conn.execute(sqlxx)
					sqlB=" if exists (select * from empfileB   where empid='"& empid&"' )     "&_
						"update EMPFILEB  set  b_whsno='"& whsno &"' ,b_groupid='"& groupid &"', b_zuno= '"& zuno &"', b_shift= '"& shift &"', "&_
						"b_job = '"& job &"' , mdtm=getdate(), muser='"&SESSION("NETUSER") &"'  WHERE EMPID='"& EMPID &"' "&_
						"else insert into empfileB ( empid,   mdtm, muser , b_whsno, b_groupid, b_zuno, b_shift, b_job ) "&_ 
						"values ( '"& EMPID &"' ,   getdate(), '"& SESSION("NETUSER") &"','"& whsno &"' ,'"& groupid &"','"& zuno &"','"& shift &"'  , '"& job &"') " 										
					conn.execute sqlB
				end if  
				
				sqlxy="update bempj set whsno='"& whsno &"' where empid='"& empid &"' and yymm='"& yymm &"' "
				conn.execute(sqlxy)
				
				sqlT1="SELECT * FROM BASICCODE WHERE  FUNC='whsno' and sys_type='"& whsno &"' "
				Set rsTmp1 = Server.CreateObject("ADODB.Recordset")
				rsTmp1.open sqlT1, conn, 1, 3				
				tmpRec(i, j, 28) = rsTmp1("sys_value")
				
				sqlT2="SELECT * FROM BASICCODE WHERE  FUNC='groupid' and sys_type='"& groupid &"' "
				Set rsTmp2 = Server.CreateObject("ADODB.Recordset")
				rsTmp2.open sqlT2, conn, 1, 3
				tmpRec(i, j, 29) = rsTmp2("sys_value")
				
				if trim(zuno)<>"" then 
					sqlT3="SELECT * FROM BASICCODE WHERE  FUNC='zuno' and sys_type='"& zuno &"' "
					Set rsTmp3 = Server.CreateObject("ADODB.Recordset")
					rsTmp3.open sqlT3, conn, 1, 3
					tmpRec(i, j, 30) = rsTmp3("sys_value")
				else
					tmpRec(i, j, 30)=""
				end if	
				
				
				sqlT4="SELECT * FROM BASICCODE WHERE  FUNC='shift' and sys_type='"& shift &"' "
				Set rsTmp4 = Server.CreateObject("ADODB.Recordset")
				rsTmp4.open sqlT4, conn, 1, 3
				tmpRec(i, j, 31) = rsTmp4("sys_value")
				
				
			
				str1=" → " &tmpRec(i, j, 28) &"<BR>"
				str2=" → " &tmpRec(i, j, 29)&" "&trim(tmpRec(i, j, 30)) &"<BR>"
				str3=" → " &tmpRec(i, j, 31) 
				x=x+1
				y=y& X &". "&trim(tmpRec(i, j, 13))&" "&empid&" "&empinfo& "<BR>異動日期: " & tmpRec(i, j, 25)  &"<br>異動內容:<BR>" & str1 & str2 & str3 &"<BR>異動說明:<BR>" & trim(memo) &"<hr size=0	style='border: 1px dotted #999999;' align=left >"
				
				set rsTmp1=nothing
				set rsTmp2=nothing
				set rsTmp3=nothing
				set rsTmp4=nothing 
			end if	
		end if 		
	next
next 
'response.write Y 
'response.end 
if x=0 then 
	'conn.close
	set conn=nothing
	response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td align=center>"
	response.write "<font class=txt12 color=red>無修改資料!!</font><BR><BR>"	
	response.write "</td></tr>"
	response.write "<tr><td align=center><BR><BR><a href='YEBBB0202.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
	response.write "</td></tr></table>"
	response.end 
else
	if  conn.errors.count=0 or err.number=0   then 
		conn.CommitTrans
		conn.close
		set conn=nothing
		set Session("YEBBB0102")=nothing 
		'response.redirect "yeaae0102.Fore.asp?A1="&A1
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"
		response.write "<font class=txt12 color=red>資料處理成功OK(SUCCESS)!!</font><BR><BR>"
		response.write y
		response.write "</td></tr>"
		response.write "<tr><td align=center><BR><BR><a href='YEBBB0202.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"
		response.end 
		
	else	
		conn.close
		set conn=nothing
		response.write "errors!!"
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"	
		response.write "</td></tr>"
		response.write "<tr><td><BR><BR><a href='YEBBB0202.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"	
		Response.End 
	end if 	 
end if 	
%>
</body>
</html>