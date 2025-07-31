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
 
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")
yymm = request("yymm") 
dat1 = request("Dat1")

TYM = left(dat1,4)&mid(dat1,6,2)

'response.write yymm &"<BR>"

Set conn = GetSQLServerConnection()
tmpRec = Session("YEBBB0201") 

nowYm=year(date())&right("00"&month(date()),2)

 

conn.BeginTrans
x=0
y=""
for i = 1 to TotalPage 
	for j = 1 to PageRec 
		'response.write tmpRec(i, j , 0) &"<BR>"
		empinfo=trim(tmpRec(i, j, 2))&" "&trim(tmpRec(i, j, 3)) 
		if tmpRec(i, j , 0) = "del" then		
			sql="delete  basicCode where  autoid='"& trim(tmpRec(i, j, 4)) &"' "	
			conn.execute(sql)	
		else
			if trim(request("op")(j))="upd" or yymm<>nowYm then 
				empid=trim(tmpRec(i, j, 1))
				country=trim(tmpRec(i, j, 4))
				whsno=trim(request("Fwhsno")(j))				
				groupid=trim(request("Fgroupid")(j))				
				zuno=trim(request("Fzuno")(j))								
				shift=trim(request("Fshift")(j))
				job=trim(tmpRec(i, j, 6))
				memo=REPLACE(trim(request("memo")(j)),vbCrLf,"<BR>")
				
				
				sqlC="select * from bempg where yymm ='"& yymm &"' and empid='"& empid &"' order by yymm desc , aid desc  " 
				'response.write sqlc&"<BR>"
				Set rst2 = Server.CreateObject("ADODB.Recordset")
				rst2.open sqlC, conn, 3,3  
				if rst2.eof then  
					sqlz="insert into BempG (yymm, empid, whsno, country, groupid, zuno, shift, memo, mdtm, muser ) values ( "&_
						 "'"& yymm &"','"& empid &"','"& whsno &"','"& country &"','"& groupid &"','"& zuno &"','"& shift &"','', "&_
						 "getdate(), '"& session("NETuser") &"' ) " 
					conn.execute(sqlz) 		
					'response.write sqlz &"<BR>"			
				else 
					sqlz="update BempG set  whsno='"&whsno&"', groupid='"&groupid&"', zuno='"&zuno&"',"&_
						 "shift='"& shift &"',   memo='"& memo &"',mdtm=getdate() , muser='"& session("NETuser") &"' "&_
						 "where empid='"& empid &"' and yymm = '"& yymm  &"' "	
					conn.execute(sqlz)	 
					'response.write sqlz &"<BR>"
				end if	 
				
				
				sqlT1="SELECT * FROM BASICCODE WHERE  FUNC='whsno' and sys_type='"& whsno &"' "
				set rsTmp1=conn.execute(SqlT1)
				tmpRec(i, j, 25) = rsTmp1("sys_value")
				set rstmp1=nothing
				sqlT2="SELECT * FROM BASICCODE WHERE  FUNC='groupid' and sys_type='"& groupid &"' "
				set rsTmp2=conn.execute(SqlT2)
				tmpRec(i, j, 26) = rsTmp2("sys_value")
				set rstmp2=nothing
				
				if trim(zuno)<>"" then 
					sqlT3="SELECT * FROM BASICCODE WHERE  FUNC='zuno' and sys_type='"& zuno &"' "
					set rsTmp3=conn.execute(SqlT3)
					tmpRec(i, j, 27) = rsTmp3("sys_value")
				else
					tmpRec(i, j, 27)=""
				end if	
				set rstmp3=nothing
				
				if trim(tmpRec(i, j, 21))="ALL" then 
					sstr="日"
				elseif trim(tmpRec(i, j, 21))="A" then 
					sstr="A班"	
				elseif trim(tmpRec(i, j, 21))="B" then 
					sstr="B班"
				else
					sstr=""
				end if 			
				
				sqlT4="SELECT * FROM BASICCODE WHERE  FUNC='shift' and sys_type='"& shift &"' "
				set rsTmp4=conn.execute(SqlT4)
				tmpRec(i, j, 28) = rsTmp4("sys_value")
			
				str1=trim(tmpRec(i, j, 11))&" → " &tmpRec(i, j, 25) &"<BR>"
				str2=trim(tmpRec(i, j, 13))&" "&trim(tmpRec(i, j, 14)) &" → " &tmpRec(i, j, 26)&" "&trim(tmpRec(i, j, 27)) &"<BR>"
				str3=trim(sstr)&" → " & tmpRec(i, j, 28)
				x=x+1
				y=y& X &". "&trim(tmpRec(i, j, 13))&" "&empid&" "&empinfo& "<BR>處理年月: " & yymm &"<br>異動內容:<BR>" & str1 & str2 & str3 &"<BR>異動說明:<BR>" & trim(memo) &"<hr size=0	style='border: 1px dotted #999999;' align=left >"
				
			end if	
		end if 		
	next
next 

'response.write Y 
'response.end 
Session("YEBBB0201") = tmpRec

if x=0 then 
	conn.close
	set conn=nothing
	response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td align=center>"
	response.write "<font class=txt12 color=red>無修改資料!!</font><BR><BR>"	
	response.write "</td></tr>"
	response.write "<tr><td align=center><BR><BR><a href='YEBBB0201.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
	response.write "</td></tr></table>"
	response.end 
else
	if  conn.errors.count=0  then 
		conn.CommitTrans
		conn.close
		set conn=nothing
		'response.redirect "yeaae0101.Fore.asp?A1="&A1		
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"
		response.write "<font class=txt12 color=red>資料處理成功OK(SUCCESS)!!</font><BR><BR>"
		response.write y
		response.write "</td></tr>"
		response.write "<tr><td align=center><BR><BR><a href='YEBBB0201.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"
		set session("YEBBB0201")=nothing
		response.end 
		
	else	
		conn.close
		set conn=nothing
		response.write "errors!!"
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"	
		response.write "</td></tr>"
		response.write "<tr><td><BR><BR><a href='YEBBB0201.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"	
		Response.End 
	end if 	 
end if 	
%>
</body>
</html>