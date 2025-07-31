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


d1 = request("d1")
d2 = request("d2")
tim1 = request("tim1")
tim2 = request("tim2")
whour= request("whour")
studyName= request("studyName")
ssno = request("ssno")
studyGroup = request("studyGroup")
teacher = request("teacher")
nw = request("nw")
amt = request("amt")
dm = request("dm")
memo=request("memo")
pzj=request("pzj")

planyear = left(d1,4)
dim T(12)
for zz=1 to 12
	if zz=int(mid(d1,6,2)) then
		T(zz)="Y"
	else
		T(zz)=""
	end if 	
	'response.write  zz & T(zz) & "<BR>"
next 

Set conn = GetSQLServerConnection()  

flag = request("flag")
if flag="del" then  
	sqlx="update  empstudy set status='D' where ssno='"& ssno &"' and convert(char(10),d1,111)='"& d1 &"' and convert(char(10),d2,111)='"& d2 &"' "
	conn.execute(sqlx)
	'response.redirect "ysbe0503.asp"

%>	<script language=vbs > 		
		open "yebe0503.asp" , "Fore"
	</script>
<%	response.end 
end if 
conn.BeginTrans

x=0 
y="" 
for i = 1 to 20 
	op = request("op")(i)
	aid = request("aid")(i)
	empid = request("empid")(i)
	groupid = request("groupid")(i)
	whsno = request("whsno")(i)
	country = request("country")(i)
	pzjno= request("pzjno")(i)
	'response.write i &"-"& op &"<Br>"	
	if empid<>"" then 
		if op="DEL" then 	
			sqln="update  empstudy set status='D' where empid='"& empid &"' and aid='"& aid &"' "
			conn.execute(Sqln)	 
			'response.write sqln &"<Br>"
			x = x + 1 
		else		
			if aid="" then 
				sqln="insert into empstudy( whsno,ssno,empid,groupid,country,D1,D2,Tim1,Tim2,Whour,StudyName,StudyGroup,Teacher,status,memo,NW,mdtm,muser,pzjno) values ( "&_
					 "'"& whsno &"','"& ssno &"', '"& empid &"', '"& groupid &"', '"& country &"',  '"& D1 &"', '"& D2 &"', '"& Tim1 &"', '"& Tim2 &"', "&_
					 "'"& Whour &"',N'"& StudyName &"',N'"& StudyGroup &"',N'"& Teacher &"','','"& memo &"', '"& nw &"',getdate(), "&_
					 "'"& session("NETUSER") &"','"& pzjno &"') "
				conn.execute(Sqln)	 
				'response.write sqln &"<Br>"
				x = x + 1 
			else
				sqln="update empstudy set studyname=N'"& StudyName &"', studygroup=N'"& StudyGroup &"', teacher=N'"& Teacher&"', "&_
					 "pzjno='"& pzjno &"',mdtm=getdate(), muser='"& session("NETUSER") &"' "&_
					 "where empid='"& empid &"' and aid='"& aid &"' and ssno='"& ssno &"'"
				conn.execute(Sqln)	 
				'response.write sqln &"<Br>" 
				x = x + 1  
			end if 	
		end if 
	end if 	
next 
 
'response.write Y 
'response.end 
if x=0 then 
	conn.close
	set conn=nothing
	response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td align=center>"
	response.write "<font class=txt12 color=red>無修改資料!!</font><BR><BR>"	
	response.write "</td></tr>"
	response.write "<tr><td align=center><BR><BR><a href='yebe0503.ASP?totalpage=0&planyear='"& planyear &"''><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
	response.write "</td></tr></table>"
	response.end 
else
	if  err.number=0  then 
		conn.CommitTrans
		conn.close
		set conn=nothing
		'response.redirect "yebe0501.Fore.asp?planyear="&planyear		
		response.write "<table width=550 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"
		response.write "<font class=txt12 color=red>資料處理成功OK(SUCCESS)!!</font><BR><BR>"
		response.write y
		response.write "</td></tr>"
		response.write "<tr><td align=center><BR><BR><a href='yebe0503.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"
		set session("YEBE0501")=nothing
		response.end 
		
	else
		conn.close
		set conn=nothing	
		response.write "errors!!"
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td>"	
		response.write "</td></tr>"
		response.write "<tr><td><BR><BR><a href='yebe0503.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"	
		Response.End 
	end if 	 
end if 	
%>
</body>
</html>