<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="yebbb03"   
nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date()-1)&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if    
empid=ucase(trim(request("empid")))
rpno = request("rpno")
rptype = request("rptype")

sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
	"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where empid='"& empid &"' and rpno='"& rpno &"' and rp_type='"& rptype &"' ) a "&_
	"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
	"left join ( select *from basicCode ) c on c.func=case when a.rp_type='R' then 'goods' else case when a.rp_type='P' then 'bads' else '' end end  and c.sys_type = a.rp_func "
	'response.write sql 
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 1, 3 
if not rs.eof then
	rpwhsno = rs("rpwhsno")
	rp_type=rs("rp_type")
	rp_func=rs("rp_func")
	rp_dat = rs("rp_date")
	rp_method=rs("rp_method")
	rp_memo=replace(replace(rs("rpmemo"),"<BR>", vbCrLf),"<br>",vbCrLf)
	empname=rs("empnam_cn")&" "&rs("empnam_vn")
	indat=rs("nindat")
	gstr = rs("gstr")&"-"&rs("zstr")
	jstr = rs("jstr")&" "
	whsno = rs("whsno")	
	autoid = rs("autoid")	
	fileno = rs("fileno")
end if 
	
	'response.write fileno
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT  LANGUAGE=javascript>
 
function f(){
	<%=self%>.rp_dat.focus();
}  
 

function empidchg(){
	if(<%=self%>.empid.value !=""){ 
	 	<%=self%>.action = "<%=self%>.ForeGnd.asp";
	 	<%=self%>.submit();
	} 	
}
 
</SCRIPT>   
</head> 
<body  onload='f()' >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=HIDDEN NAME="act" VALUE="">
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="94%" align="center" >
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3> 		 
					<tr>
						<Td width=90 align=right>No.</td>
						<td>
							<input type="text" style="width:150px" name="rpno" size=8  readonly class="readonly" value="<%=rpno%>" >
							<input type=hidden name="aid"   value="<%=autoid%>" >
						</td>
						<td>
							&nbsp;&nbsp;&nbsp;
							<font class=txt style="cursor:pointer" onclick="history.back()" color=blue>返回主畫面(Back)</font>				
						</td>
					</tr>	
					<tr>
						<Td width=90 align=right>工號/姓名</td>
						<td>
							<input type="text" name="empid" size=8  readonly class="readonly" value="<%=empid%>"  >
						</td>
						<td>
							<input type="text" name="empname" size=30  readonly class="readonly" value="<%=empname%>" style='font-size:8pt'>
						</td>									
					</tr>				 
					<tr>
						<Td align=right>單位/職務</td>
						<td>
							<input type=hidden name="whsno" value="<%=whsno%>">
							<input type="text" name="gstr" size=20  readonly class="readonly" value="<%=gstr%>">
						</td>
						<td>
							<input type="text" name="jstr" size=18  readonly class="readonly" value="<%=jstr%>">
						</td>
						<Td align=right>到職日</td>
						<td>
							<input type="text" style="width:150px" name="indat" size=15  readonly class="readonly" value="<%=indat%>">				
						</td>
					</tr>			
					<tr>
						<Td align=right>獎懲類別</td>
						<td>
							<select name=rp_type    disabled class="readonly">
								<option value="" <%if rp_type="" then%>selected<%end if%>>-----</option>
								<option value="R" <%if rp_type="R" then%>selected<%end if%>>R-獎</option>
								<option value="P" <%if rp_type="P" then%>selected<%end if%>>P-懲</option>
							</select>
						</td>
						<td>
							<select name=rp_func  onchange='empidchg()'>
								<%if rp_type="" then%>
									<option value="">-----</option>
								<%else%>
									<%
									if rp_type="R" then 
										sql="select* from basicCode  where func='goods' order by sys_type"
									elseif 	rp_type="P" then 
										sql="select* from basicCode  where func='bads' order by sys_type"
									end if 	
									SET RST=COnn.execute(sql)
									while not rst.eof 
									%>
									<option value="<%=rst("sys_type")%>" <%if trim(rp_func)=trim(rst("sys_type")) then %>selected<%end if%>><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
									<%rst.movenext%>
									<%wend
									rst.close
									set rst=nothing
								end if 
								conn.close 
								set conn=nothing
									%> 
							</select>
						</td>
						<Td align=right>事件日期</td>
						<td>
							<input type="text" name="rp_dat" id="rp_dat" size=15  value="<%=rp_dat%>" onblur="chkdat('rp_dat')">
						</td>
					</tr>				 		
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table id="myTableForm" width="100%">
					<tr height="40px">
						<td>&nbsp;&nbsp;處理方式說明</td>
					</tr>
					<tr align="center">
						<td align=center><input type="text" style="width:98%" name="rp_method" value="<%=rp_method%>"> </td>
					</tr>
					<tr height="30px">
						<td>&nbsp;&nbsp;事件內容說明</td>
					</tr>
					<tr>
						<td align=center>
							<TEXTAREA rows=10 cols=75 name=rp_memo  STYLE='HEIGHT:AUTO' wrap="PHYSICAL"><%=rp_memo%></TEXTAREA>				
						</td>
					</tr>	
					<tr height="30px">
						<td>&nbsp;&nbsp;文件編號</td>
					</tr>
					<tr>
						<td align=center>
							<input type="text" style="width:98%" name="fileno" value="<%=fileno%>" >
						</td>
					</tr>			
					<Tr>
						<Td align=center height="50px">				
							<INPUT TYPE=button name=btn value="(Y)Confirm" onclick=go() class="btn btn-sm btn-danger">
							<INPUT TYPE=button name=btn value="(N)Cancel" onclick=goclr() class="btn btn-sm btn-outline-secondary">
							&nbsp;								
							<INPUT TYPE=button name=btn value="(M)回主畫面" onclick="history.back()" class="btn btn-sm btn-outline-secondary">
							&nbsp;				
							&nbsp;
							<INPUT TYPE=button name=btn value="(D)刪除資料" onclick="godel()" class="btn btn-sm btn-outline-secondary" style='background-color:#ffcccc' >
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
			 
</body>
</html>


<script language=javascript>
function goclr(){
	<%=self%>.empid.value="";
	<%=self%>.action = "<%=self%>.new.asp";
	<%=self%>.submit();
}  
 
	
function go(){
	if(<%=self%>.empid.value.trim()==""){ 
		alert( "請輸入員工編號!!");
		<%=self%>.empid.focus();
	}else if(<%=self%>.rp_dat.value.trim()==""){ 
		alert( "請輸入事件日期!!");
		<%=self%>.rp_dat.focus();
	}else if(<%=self%>.rp_type.value.trim()==""){ 
		alert( "請輸入獎懲類別!!");
		<%=self%>.rp_type.focus();
	}else if(<%=self%>.rp_func.value.trim()==""){ 
		alert( "請輸入獎懲方式!!");
		<%=self%>.rp_func.focus();
	}else if(<%=self%>.rp_memo.value.trim()==""){ 
		alert( "請輸入事件內容說明!!");
		<%=self%>.rp_memo.focus();
	}else{			
		<%=self%>.act.value="upd";	
		<%=self%>.action="<%=self%>.upd.asp";
		<%=self%>.submit();
	}	
}  

function godel(){
	if (confirm("確定要刪除此筆資料?",64)){ 
		<%=self%>.act.value="del";
	 	<%=self%>.action="<%=self%>.upd.asp";
	 	<%=self%>.submit() ;
	} 	
} 	


</script> 