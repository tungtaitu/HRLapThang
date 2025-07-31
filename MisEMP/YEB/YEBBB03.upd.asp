<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="yebbb03"  

if session("netuser")="" then 
	errmsg="使用者帳號為空,請重新登入!!"
	goerr()
end if 

empid=request("empid") 
rp_dat = request("rp_dat")   
rp_type = request("rp_type")
rp_func = request("rp_func")
rp_method = request("rp_method")
rp_memo = request("rp_memo") 
rp_memo=trim(REPLACE (rp_memo,vbCrLf,"<br>")) 
rpno = trim(request("rpno"))
nowmonth =trim(request("nowmonth"))
whsno = request("whsno")
act = request("act")
autoid = request("aid") 
fileno = request("fileno") 

conn.BeginTrans    

if rp_dat="" then 
	yymmstr=nowmonth 
else
	yymmstr = left(replace(rp_dat,"/",""),6)
end if 	

if act="del" then 
	sql="update emprepe set status='D', mdtm=getdate(), muser='"&session("netuser")&"' where autoid='"& autoid &"' and rpno='"& rpno &"' "
	conn.execute(sql)
elseif act="upd" then 
	sql="update emprepe set rp_dat='"& rp_dat&"', rp_method=N'"& rp_method &"', rpmemo=N'"& rp_memo &"', "&_
		"rp_func='"&rp_func&"', mdtm=getdate(), muser='"&session("netuser")&"', fileno='"& fileno &"'  "&_
		"where autoid='"& autoid &"' and rpno='"& rpno &"' "
	conn.execute(sql)
else 
	if rpno = "" then 
		Set rs = Server.CreateObject("ADODB.Recordset")
		sqln="exec GS_GetSSno 'rpno', '"& yymmstr &"' , '"& whsno &"', '' , '' "
		'response.write sqln&"<RB>"
		rs.open sqln, conn, 1, 3
		if rs("msg")="" then 
			pno = rs("pno")
			sql="insert into emprepe (rpno, rpwhsno, empid, rp_dat, rp_type, rp_func, rp_method, rpmemo, fileno, mdtm, muser ) values ( "&_
				"'"& pno &"','"& whsno &"', '"&empid&"','"&rp_dat&"','"&rp_type&"','"&rp_func&"',N'"&rp_method&"',N'"&rp_memo&"',"&_
				"'"& fileno &"',getdate(), '"&session("netuser")&"' ) "
			'response.write sql &"<RB>"	
			conn.execute(Sql) 
			
		else
			errmsg="取號錯誤!!資料處理失敗!!"
			goerr()
		end if 		
		set rs=nothing 
	end if  
end if 	
'response.write sql 
'response.end  
session("rpwhsno") = whsno 
function goerr()
	response.write errmsg
	conn.close 
	set conn=nothing
	response.end 
end function 
if err.number = 0 then
	conn.CommitTrans  
	conn.close 
	set conn=nothing
	if act="" then 
		msg="編號: " & rp_type&pno & ", 工號so the: " & empid &"  新增成功!!" &"<BR><BR><BR>"
		response.write "<table width=500><tr><td>"
		response.write "<center><BR>"
		response.write msg 
		response.write "<a href=yebbb03.new.asp target=_self>繼續新增(Add New)</a><BR>"
		response.write "<a href=yebbb03.fore.asp?whsno="&whsno&"&rpno="&rpno&"target=_self>返回主畫面(BACK)</a>"
		response.write "</center>"
		response.write "</td></tr></table>"
		response.end 
	else	
%>	<script language=javascript>
		F_rpno = "<%=rpno%>";
		alert( "資料處理成功(OK)!!");
		open( "<%=self%>.Fore.asp?whsno="+"<%=whsno%>"+"&rpno="+F_rpno , "_self")
	</script>
<%	end if
else
	conn.RollbackTrans
	conn.close 
	set conn=nothing
%>	<script language=javascript>
		alert( "資料處理失敗!!");
		open( "<%=self%>.asp", "_self");
	</script>	
<%end if%>	
 