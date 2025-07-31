<%@Language=VBScript codepage=65001%>
<!-------- #include file ="../GetSQLServerConnection.fun" --------->
<!--#include file="../ADOINC.inc"-->
<!--#include file="../include/sideinfo.inc"-->
<%

Set conn = GetSQLServerConnection()
'Set rs = Server.CreateObject("ADODB.Recordset")

self="YEAAE0201"  
'response.write "???=" & session("netuserip")
'sql="select * from sysuser where muser='"& session("netuser") &"'" 
'rs.open sql, conn, 3, 3   
empid=request("empid") 
groupid="" 
jobid=""
whsno=""
muser=""
username=""
'hitstr=""
if empid <>""  then  
	sql="select isnull(convert(char(10),a.outdat,111),'') as o_dat, *   "&_
		  ", groupid=isnull( (select top 1  groupid from bempg where empid=a.empid order by yymm desc),'') "&_
			", whsno=  isnull( (select top 1  whsno   from bempg where empid=a.empid order by yymm desc),'') "&_
			", jobid = isnull( (select top 1  job from bempj where empid=a.empid order by yymm desc),'') "&_
			"from  empfile a  where empid<>'pelin' and  empid='"& trim(request.Form("empid"))&"'    "
	set rs=conn.execute(sql)
	'response.write sql 
	if not rs.eof then 
		groupid = rs("groupid")
		whsno = rs("whsno")
		jobid = rs("jobid")
		muser=rs("empid")
		if rs("empnam_cn")="" then 
			username=rs("empnam_vn")
		else	
			username=rs("empnam_cn")
		end if  
		if rs("o_dat")<>"" then 
			'response.write "axxxx"
			hitstr=rs("o_dat")&"已離職"
		else
			'response.write "bbb"
			hitstr=""
		end if 	
	end if 
end if  
'response.write hitstr 
%>
<HTML>
<HEAD>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">

<SCRIPT  LANGUAGE=javascript>
<!--
'-----------------enter to next field
function enterto(){
	if(window.event.keyCode == 13) window.event.keyCode =9;
}

function sch(){
	if(<%=self%>.empid.value !=""){ 
		<%=self%>.action="<%=self%>.fore.asp";
		<%=self%>.submit();
	}
}
-->
</SCRIPT>
</HEAD>

<body  onload=f()  onkeydown="enterto()" >	
<FORM action="<%=SELF%>.UpdateDB.asp" method="POST" name="<%=SELF%>" >
	<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" border=0 >
		<tr>
			<td align="center">
				<table id="myTableForm" width="60%">		
					<tr><td colspan=4>&nbsp;</td></tr>
					<tr>
						<td class="frmtable-label">員工編號<br>So the</td>
						<td>
							<input type="text" style="width:100px" size=15 name="empid" maxlength="6"  value="<%=empid%>"   onchange="sch()" >
							<SPAN style="color:red"><%=hitstr%></span>
						</td>
						<td class="frmtable-label">帳號<br>ID</td>
						<td align=left>
							   <input type="text" style="width:100px" size=8 name="muser" id="muser" value="<%=muser%>" maxlength="6"  >
						</td>
					</tr>
					<tr>					 
						<td class="frmtable-label">姓名<br>Ten</td>		 
						<td colspan="3"> <input type="text" size=25 name="Username" id="Username" value="<%=Username%>"  > </td>
					</tr>
					<tr>
						<td class="frmtable-label">廠別<br>Xuong</td>
						<td>
							<select name=whsno >
							<option value=""  <%IF  trim(whsno)="" then%>selected<%end if%> />----- 
							<%sql="select * from basicCode where func='WHSNO' order by sys_type"
							  set rds=conn.execute(sql)
							  while not rds.eof
							%>
								<option value="<%=rds("sys_type")%>"  <%IF  trim(whsno)=trim(rds("sys_type")) then%>selected<%end if%> ><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
							<%rds.movenext
							wend
							%>
							</select>
						</td>		
						<td class="frmtable-label">單位<br>bo phan</td>
						<td>
							<select name=groupid >
							<option value=""  <%IF  trim(groupid)="" then%>selected<%end if%> />----- 
							<%sql="select * from basicCode where func='groupid' order by sys_type"
							  set rds=conn.execute(sql)
							  while not rds.eof
							%>
								<option value="<%=rds("sys_type")%>"  <%IF  trim(groupid)=trim(rds("sys_type")) then%>selected<%end if%>  ><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
							<%rds.movenext
							wend
							%>
							</select>
						</td>
					</tr>	
					<tr >
						<td class="frmtable-label">職等<br>Chuc vu</td>
						<td>
							<select name=jobid >
							<option value=""  <%IF  trim(jobid)="" then%>selected<%end if%> />----- 
							<%sql="select * from basicCode where func='lev' order by sys_type"
							  set rds=conn.execute(sql)
							  while not rds.eof
							%>
								<option value="<%=rds("sys_type")%>" <%IF  trim(jobid)=trim(rds("sys_type")) then%>selected<%end if%>   ><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
							<%rds.movenext
							wend 
							rds.close : set rds=nothing 
							%>
							</select>
						</td>
						<td class="frmtable-label">權限<br>Loai</td>
						<td>
							<select name=rights >
							<%
							  if session("netuser")="PELIN"    then 
								sql="select * from basicCode where func='grp' order by sys_type"
							  else
									if session("rights") <="1"    then  
										sql="select * from basicCode where func='grp' and   sys_type in ( '1','2','3', '7' , 'B', 'C')  order by sys_type"
									elseif session("rights") ="2"    then 
										sql="select * from basicCode where func='grp' and   sys_type in ( '2','3', '7' , 'B', 'C')  order by sys_type"
									elseif session("rights") ="A"    then 
										sql="select * from basicCode where func='grp' and   sys_type in (  '3', 'A' , 'B', 'C')  order by sys_type"	
									else
								sql="select * from basicCode where func='grp' and   sys_type in ( '3', '7' , 'B', 'C')  order by sys_type"
									end if 
							  end if	
							  set rds=conn.execute(sql)
							  while not rds.eof
							%>
								<option value="<%=rds("sys_type")%>"   ><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
							<%rds.movenext
							wend 
							rds.close : set rds=nothing 
							%>
							</select>
						</td>
					</tr>				
					<tr>
						<td class="frmtable-label">密碼<br>mat ma</td>
						<td>
							<input type="password" size=22 name="Password" id="Password" maxlength="10" style="font size: 9pt" ><br>
						</td>		
						<td class="frmtable-label">確認密碼<br>mat ma</td>
						<td >
							<input type="password" size=22 name="Password2" id="Password2" maxlength="10" style="font size: 9pt" ><br>
						</td>
					</tr>				
					<tr > 
						<td align="CENTER"  colspan=4 height="50px">
							<%if UCASE(session("mode"))="W" then%>
								<input type="button" name="send" id="send" value="(Y)Confirm" onclick="go()" class="btn btn-sm btn-danger">
								<input type="button" name="btcancel" id="btcancel" value="(N)Cancel" class="btn btn-sm btn-outline-secondary" onclick="clr()">
							<%end if%>				
						</td>
					</tr>
				</table>
						
			</td>
		</tr>
	</table>
	
 </FORM>

<%
conn.close : set conn=nothing 
%>
</BODY>
</HTML>

<script   language="javascript">
	function clr()
	{
		open("<%=self%>.asp" ,"Fore");
	}

	function go()
	{
		
		if(<%=self%>.muser.value=="" || <%=self%>.Username.value=="")
		{	
			alert("請輸入使用者帳號姓名!!");
			<%=self%>.muser.focus();		
		} 	
		else if(<%=self%>.Password.value =="" || (<%=self%>.Password.value != <%=self%>.Password2.value))
		{
			   alert("密碼有誤或不得為空白!!請重新輸入!!");
			   <%=self%>.Password2.value = "";	
		}   
		else
		{
				<%=self%>.submit();
		}
	}
	
	function f()
	{
		if(<%=self%>.empid.value=="")
		{ 
			<%=self%>.empid.focus();
		}
		else
		{
			<%=self%>.muser.focus();
		} 
		 
	}
	
	/*
	function chg1()
	{
		<%=self%>.muser.value=Ucase(<%=self%>.muser.value)
	}
	*/
</script>


