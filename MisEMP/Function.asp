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

<SCRIPT TYPE="text/javascript">

    function clikker(a, b) {
        if (a.style.display == '') {
            a.style.display = 'none';
            b.src = 'picture/dot_y2.gif';
        }
        else {
            a.style.display = '';
            b.src = 'picture/dot_y3.gif';
        }
    }
	
</SCRIPT>

</head>
 
<body leftmargin="1" topmargin="0" marginwidth="0" marginheight="0">
	  	  
   <table width="100%" border="0" cellspacing="3" cellpadding="1" bgcolor="#ffffff" class="txt">
	   <tr >
			<td align="left" colspan="2"><span class="toplang"><%=tit%></span></td>
	   </tr>
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
		<tr>
			
			<td align="left" colspan="2" nowrap>
				<i class="fa fa-book"></i>&nbsp;<span class="toplang"><%=left(program_id,1)&C%></span>
			</td>	         
		</tr>
				
	   <%	       
			end if
			while not rs2.eof 
	   %>	       
				<tr>
					
					<td valign="top"><i class="fa fa-mail-forward fa-flip-vertical"></i></td>
					<td align="left" nowrap>	 			   
					<%if rs2("layer") = 0 then %>
													   
						   <%if Trim(rs2("virtual_path"))="" then  %>							       
								<span class="toplang"><%=rs2("program_name")%></span>
						   <%else%>
								<a class="menuleft" href="<%=rs2("virtual_path")%>?pgid=<%=rs2("program_id")%>&pgname=<%=server.urlEnCode(rs2("program_name"))%><%=server.urlEnCode(rs2("proname_vn"))%>"  target="main" >
									<span class="toplang"><%=rs2("program_name")%></span><br>
									<span class="bottomlang"><%=LCASE(rs2("proname_vn"))%></span>
								</a>
						   <%end if%>
					   
					<%else%>							   
						<div id="<%=rs2("program_id")%>" onclick="clikker(<%=rs2("program_id")&"1"%>,<%=rs2("program_id")&"P"%>);event.returnValue=0;event.cancelBubble = true;">												
							
							<a href="#" class="menuleft" id="<%=rs2("program_id")&"P"%>" >
								<span class="toplang"><%=rs2("program_name")%></span><br>
								<span class="bottomlang"><%=LCASE(rs2("proname_vn"))%></span>									
							</a>									
						</div>
						<div id="<%=rs2("program_id")&"1"%>" style="DISPLAY: none" onclick="window.event.cancelBubble = true;">
						<%
							sql = "select * from sysprogram where layer_up = '"& rs2("program_id") &"' and  group_r like '%"& session("LoginType") &"%' order by program_name "					
							'sql = "select * from sysprogram where layer_up = '"& rs2("program_id") &"'  "					
							'response.write sql
							set rs3 = conn.execute(sql)                                   
							while not rs3.eof 
						%>
							<table border=0 cellspacing="3" cellpadding="1"  class="txt">
								<tr>
									<td valign="top"><i class="fa fa-address-book"></i></td>
									<td>										
								<%if Trim(rs3("virtual_path"))="" then%>
										<span class="toplang"><%=rs3("program_name")%></span>
								<%else%>
										<a class="menuleft" href="<%=rs3("virtual_path")%>?pgid=<%=rs3("program_id")%>&pgname=<%=server.urlEnCode(rs3("program_name"))%><%=server.urlEnCode(rs3("proname_vn"))%>" target="main" >
											<span class="toplang"><%=rs3("program_name")%></span><br>
											<span class="bottomlang"><%=LCASE(rs3("proname_vn"))%></span>													
										</a>
								<%end if%>
									</td>
								</tr>
							</table>
						<%
							rs3.movenext
							wend
						%>
						</div>
					  <%end if%>
				   </td>
				</tr>
	   <%	rs2.movenext
			wend
	   %>		           	       
		
	   <%next%>
			  
	</table>
</body>		