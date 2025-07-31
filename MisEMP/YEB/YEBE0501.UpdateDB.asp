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
planyear = request("planyear") 




'response.write gTotalPage &"<BR>"

Set conn = GetSQLServerConnection()
tmpRec = Session("YEBE0501")

conn.BeginTrans
x=0
y=""
for i = 1 to gTotalPage 
	for j = 1 to PageRec 
		'response.write tmpRec(i, j , 0) &"<BR>"
		empinfo=trim(tmpRec(i, j, 2))&" "&trim(tmpRec(i, j, 3)) 
		yy=trim(tmpRec(i, j, 1))
		ssno=trim(tmpRec(i, j, 2))
		studyName=trim(tmpRec(i, j, 3))
		T1=trim(tmpRec(i, j, 4))
		T2=trim(tmpRec(i, j, 5))
		T3=trim(tmpRec(i, j, 6))
		T4=trim(tmpRec(i, j, 7))
		T5=trim(tmpRec(i, j, 8))
		T6=trim(tmpRec(i, j, 9))
		T7=trim(tmpRec(i, j, 10))
		T8=trim(tmpRec(i, j, 11))
		T9=trim(tmpRec(i, j, 12))
		T10=trim(tmpRec(i, j, 13))
		T11=trim(tmpRec(i, j, 14))
		T12=trim(tmpRec(i, j, 15))
		pcnt=trim(tmpRec(i, j, 16))
		hhour=trim(tmpRec(i, j, 17))
		amt=trim(tmpRec(i, j, 18))
		dm=trim(tmpRec(i, j, 19))
		nw=trim(tmpRec(i, j, 20))
		memo=trim(tmpRec(i, j, 21))
		aid=trim(tmpRec(i, j, 22))
		if amt="" then amt=0
		
		if tmpRec(i, j , 0) = "del" then		
			'sql="delete  basicCode where  autoid='"& trim(tmpRec(i, j, 4)) &"' "	
			'conn.execute(sql)	
		else
			if trim(ssno)<>"" then 
			 	sql="update studyPlan set studyName=N'"& studyName &"', T1='"&T1&"', T2='"&T2&"' ,  "&_
			 		"T3='"&T3&"', T4='"&T4&"',T5='"&T5&"', T6='"&T6&"',T7='"&T7&"', T8='"&T8&"', "&_
			 		"T9='"&T9&"', T10='"&T10&"',T11='"&T11&"', T12='"&T12&"', pcnt='"& pcnt &"', "&_
			 		"hhour='"&hhour&"', amt='"&amt&"', dm='"&dm&"', nw='"&nw&"', "&_
			 		"memo='"&memo&"' where ssno='"& ssno &"' and aid='"& aid &"' " 
			 	response.write sql &"<BR>"	
			 else
			 	if trim(yy)<>"" and trim(studyName)<>"" then 
			 		sqlx="exec GS_GetSSno 'Getssno', '"& planyear&"','' , '' "
			 		set rst=conn.execute(Sqlx)
			 		if rst("msg")="" then 
			 			pno = rst("pno")			 					 			
			 		end if	
			 		sql="insert into studyPlan (yy, ssno,studyName,t1, t2, t3, t4, t5, t6,t7, t8, t9, t10,t11,t12, amt, dm, pcnt, hhour, nw, memo ) valeus ( "&_
			 			"'"& planyear &"', '"& pno &"', '"& studyName &"', '"& T1 &"','"& T2 &"','"& T3 &"','"& T4 &"','"& T5 &"','"& T6 &"', "&_
			 			"'"& T7 &"','"& T8 &"','"& T9 &"','"& T10 &"','"& T11 &"','"& T12 &"','"& amt &"','"& dm &"','"& pcnt &"','"& hhour &"', "&_
			 			"'"& nw &"','"& memo &"' )" 
			 		response.write sql &"<BR>" 	
			 	end if 
			 end if 		
		end if 		
	next
next 
 
response.write Y 
response.end 
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
		'response.redirect "yeaae0101.Fore.asp?A1="&A1		
		conn.close
		set conn=nothing
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