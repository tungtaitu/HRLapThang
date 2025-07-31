<%@Language=VBScript Codepage=65001 %>
<!--#include file="ADOINC.inc"-->
<!--#include file="GetSQLServerConnection.fun"-->

<%
if session("netuser")="" then 
	'response.redirect  "default.asp"
	response.write "<script>"
	'response.write "window.open('../default.asp','_top');"
	response.write "window.open('default.asp','_top');"
	response.write "</script>"
end if 
set conn = GetSQLServerConnection()
program_id = request("program_id")
functionN = request("functionN")
if program_id=empty then program_id="A"
if functionN=empty then functionN="S"

'response.Write program_id
'response.Write functionN


select case functionN 
	case "S" 
		srcpage = "Function_S.asp"
	case "D" 
		srcpage = "Function_D.asp"	
	case "C"
		srcpage = "Function_C.asp"	
	case else 
		srcpage = "Function_S.asp"	
end select 	
'Response.Write program_id
sql= "select * from sysprogram where left(program_id,1) like '"& program_id &"'+'%'  order by program_id "
'Response.Write sql
'Response.End 
set rs=conn.execute(sql)
if not rs.eof then
	tit = rs("program_ID")&"."&rs("program_name")
else
	tit= ""	
end if	

'Response.Write "LoginType="&session("LoginType")
'Response.Write "tit="&tit
'response.end
%>

<html>
<head>
<title>function</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<meta name="viewport" content="width=device-width, initial-scale=1">
<script src="template/js/jQuery-3.3.1.js"></script>
<script src="template/bootstrap/js/bootstrap.min.js"></script>
<script src="template/js/sidebar-menu.js"></script>
<script src="template/datepicker/bootstrap-datepicker.js"></script>
<link rel="stylesheet" type="text/css" href="template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="template/font-awesome/css/font-awesome.css">
<link rel="stylesheet" type="text/css" href="template/css/mis.css">
<link rel="stylesheet" type="text/css" href="template/datepicker/datepicker.css">
</head>

<body leftmargin="1" topmargin="0" marginwidth="0" marginheight="0">

<ul class="sidebar-menu" id="myMenu" style="height:100%;overflow-x: auto !important;white-space: nowrap;background-color:#ffffff !important;">
    <li class="sidebar-header" style="padding-left:50px;"><%=tit%></li> 
	<% for p = 0 to 4
		select case p 
			case 0
				C=".B"
			case 1 
				C=".E"
			case 2 
				C=".F"	
			case 3 
				C=".Q"	 
			case 4 
				C=".P"	
		end select	
		idstr = program_id&right(C,1)					
		sql = "select * from sysprogram where left(program_id,2) = '"& idstr &"' and len(program_id)=4 and group_r like '%"& session("rights") &"%' order by program_name"					
		'Response.Write sql
		Set rs2 = Server.CreateObject("ADODB.Recordset")
		rs2.Open sql, conn, 3, 3
		if not rs2.EOF then
   %>	
		<i class="menu-item"></i><i class="fa fa-folder-open"></i>&nbsp;<span class="toplang"><%=left(program_id,1)&C%></span>			
	<%	       
		end if
		while not rs2.eof 
    %>
    <li id="" class="menu-item">
		<%if rs2("layer") = 0 then %>
			<%if Trim(rs2("virtual_path"))="" then  %>							       
				&nbsp;&nbsp;&nbsp;<i class="fa fa-file-text-o"></i>&nbsp;&nbsp;<span class="toplang"><%=rs2("program_name")%></span>
			<%else%>
				<a  href="<%=rs2("virtual_path")%>?pgid=<%=rs2("program_id")%>&pgname=<%=server.urlEnCode(rs2("program_name"))%><%=server.urlEnCode(rs2("proname_vn"))%>"  target="main" >
					<i class="fa fa-file-text-o"></i>
					<span class="toplang"><%=rs2("program_name")%></span><br>
					<span class="bottomlang"><%=LCASE(rs2("proname_vn"))%></span>
				</a>
			<%end if%>		   
		<%else%>									
				<a href="#"  id="<%=rs2("program_id")&"P"%>" >
					<i class="fa fa-file-text-o"></i>
					<span class="toplang"><%=rs2("program_name")%></span>
					<i class="fa fa-angle-left pull-right text-secondary"></i><br>
					<span class="bottomlang"><%=LCASE(rs2("proname_vn"))%></span>									
				</a>
		<ul class="sidebar-submenu">			
			<li id="" class="menu-item level-s">
			<%
				sql = "select * from sysprogram where layer_up = '"& rs2("program_id") &"' and  group_r like '%"& session("LoginType") &"%' order by program_name "					
				'response.write sql
				set rs3 = conn.execute(sql)                                   
				while not rs3.eof 
			%>
			<%if Trim(rs3("virtual_path"))="" then%>
					&nbsp;&nbsp;&nbsp;<i class="fa fa-minus-square"></i>&nbsp;&nbsp;<span class="toplang"><%=rs3("program_name")%></span>
			<%else%>
					<a class="border-left border-danger" href="<%=rs3("virtual_path")%>?pgid=<%=rs3("program_id")%>&pgname=<%=server.urlEnCode(rs3("program_name"))%><%=server.urlEnCode(rs3("proname_vn"))%>" target="main" >
						<i class="fa fa-minus-square"></i>
						<span class="toplang"><%=rs3("program_name")%></span><br>
						<span class="bottomlang"><%=LCASE(rs3("proname_vn"))%></span>													
					</a>
			<%end if%>	
			<%
				rs3.movenext
				wend
			%>
			</li>						
		</ul>
		<%end if%>
    </li>
	<%
			rs2.movenext
			wend
	next
	%>
	<script>
		//open sub menu
		$.sidebarMenu($('.sidebar-menu'))
	</script>
</ul>

