<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!--#include file="../include/sideinfo.inc"-->

<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

SELF = "user04"

Set conn = GetSQLServerConnection()
gTotalPage = 1
PageRec = 10   'number of records per page
TableRec = 25    'number of fields per record
Queryx = trim(request("Queryx"))
F_whsno = request("F_whsno") 
F_empid = request("F_empid") 
Set conn = GetSQLServerConnection() 

if request("TotalPage") = "" or request("TotalPage") = "0" then 
   	CurrentPage = 1 	   	   
 
		sql="select a.* , x1.sys_value as set_gstr, x2.sys_value as set_zstr, x3.sys_value as set_jstr, "&_
				"h.empnam_cn as hcqNam_cn , h.empnam_vn as hcqnam_vn, h.jstr as HcqJob,h.email as Hcqmail,"&_
				"J.empnam_cn as JcqNam_cn , J.empnam_vn as Jcqnam_vn, J.jstr as JcqJob,J.email as Jcqmail, "&_
				"Z.empnam_cn as ZcqNam_cn , Z.empnam_vn as Zcqnam_vn, z.jstr as ZcqJob,Z.email as Zcqmail "&_
				"from  "&_
				"(select * from yfycq where isnull(whsno,'') like '"&whsno&"%' ) a "&_
				"left join (select *  from basicCode where func='groupid' ) x1 on x1.sys_value = a.groupid "&_
				"left join (select *  from basicCode where func='zuno' ) x2 on x1.sys_value = a.zuno "&_
				"left join (select *  from basicCode where func='lev' ) x3 on x1.sys_value = a.job "&_
				"left join (select  empid, empnam_cn, empnam_vn , jstr,email from   view_empfile ) z on z.empid = a.zcqid "&_
				"left join (select  empid, empnam_cn, empnam_vn , jstr,email from   view_empfile ) j on j.empid = a.jcqid "&_
				"left join (select  empid, empnam_cn, empnam_vn , jstr,email from   view_empfile ) h on h.empid = a.hcqid "&_
				"order by a.groupid "
 
  
	'response.write sql 
	'response.end 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 1, 3
	if not rs.eof then 
		pagerec= pagerec+10		   
	   rs.PageSize = PageRec
	   RecordInDB = rs.RecordCount 
	   TotalPage = rs.PageCount 	
	   gTotalPage = totalpage
	end if  		 
	'Set conn = nothing 	
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array	
	for i = 1 to TotalPage 
	   for j = 1 to PageRec
		   if not rs.EOF then 				  
				  tmpRec(i, j, 0) = "no"
				  tmpRec(i, j, 1) = rs("whsno")
					tmpRec(i, j, 2) = rs("groupid")
					tmpRec(i, j, 3) = rs("set_gstr")
					tmpRec(i, j, 4) = rs("zuno")
				  tmpRec(i, j, 5) = rs("set_zstr") 
				  tmpRec(i, j, 6) = rs("job")
				  tmpRec(i, j, 7) = rs("set_jstr")				  
				  tmpRec(i, j, 8) = rs("zcqid")				  
				  tmpRec(i, j, 9) =  rs("zcqnam_cn")	
					tmpRec(i, j, 10) = rs("zcqnam_vn")	
					tmpRec(i, j, 11) = rs("ZcqJob")	
					tmpRec(i, j, 12) = rs("zcqmail")	
				  tmpRec(i, j, 13) = rs("jcqid")				  
				  tmpRec(i, j, 14) = rs("jcqnam_cn")	
					tmpRec(i, j, 15) = rs("jcqnam_vn")	
					tmpRec(i, j, 16) = rs("jcqJob")	
					tmpRec(i, j, 17) = rs("jcqmail")	
				  tmpRec(i, j, 18) = rs("hcqid")				  
				  tmpRec(i, j, 19) = rs("hcqnam_cn")	
					tmpRec(i, j, 20) = rs("hcqnam_vn")	
					tmpRec(i, j, 21) = rs("hcqJob")						
					tmpRec(i, j, 22) = rs("hcqmail")	
			  rs.MoveNext 
		   else 
			  exit for 
		   end if 			
	   next
	
	   if rs.EOF then 
		  rs.Close 
		  Set rs = nothing
		  exit for 
	   end if 
	next 
	Session("user04B") = tmpRec	
else
	TotalPage = cint(request("TotalPage"))	   
	gTotalPage = cint(request("gTotalPage"))
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB = cint(request("RecordInDB"))
	StoreToSession()
	tmpRec = Session("user04B") 
	
	Select case request("send") 
	      Case "FIRST"
		       CurrentPage = 1			
	      Case "BACK"
		       if cint(CurrentPage) <> 1 then 
			      CurrentPage = cint(CurrentPage) - 1				
		       end if
	      Case "NEXT"
		       if cint(CurrentPage) <= cint(gTotalPage) then 
			      CurrentPage = CurrentPage + 1 
		       end if			
	      Case "END"
		       CurrentPage = cint(TotalPage)
	      Case Else 
		       CurrentPage = 1	
	end Select 	
end if 
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
 
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<FORM NAME="<%=SELF%>" METHOD=POST ACTION="<%=self%>.fore.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 
<INPUT TYPE=hidden NAME=queryx VALUE="X"> 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table class="table-borderless table-sm text-secondary txt">
					<tr >
						<TD width=50 ALIGN=CENTER >廠別:</TD>
						<TD width=150>
							<select name=F_whsno ONCHANGE="Q_Data()" >
							<option value="" >ALL</option>
							<%sql="select * from  basiccode where func='whsno' order by sys_type"
							set rds=conn.execute(sql)
							while not rds.eof 
							%>
							<option value="<%=rds("sys_type")%>" <%IF F_whsno=rds("sys_type") THEN%> SELECTED<%END IF%>><%=rds("sys_value")%></option>
							<%rds.movenext
							wend
							set rds=nothing %> 
							</SELECT>
						</TD> 
						<TD ALIGN=CENTER >查詢:</TD>
						<TD >
							<input type="text" name=query size=20 value="<%=query%>" style='height:22'>
						</TD>
						<td>
							<input type=button name="btn" class="btn btn-sm btn-outline-secondary" value="查詢K.tra" onclick="Q_Data()">
						</td> 	
						<td width=150></td>		
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" width="98%">
					<tr height=30 class="header">
						<td align=center>刪除</td>						
						<td align=center>部門</td>
						<td align=center>組</td>				
						<td align=center >直接主管</td>
						<td align=center >間接主管</td>
						<td align=center >核決主管</td> 
					</tr>	
				  <%
					for x = 1 to PageRec			
					if x mod 2 = 0  then 
						wk_color = "#ffff99"
						
					else 
						wk_color = "#ffff99"
						
					end if  
					'Response.Write CurrentRow &"<BR>" 
					
				  %> 
					<tr>
						<td align=center>
							<input type=checkbox name=func onclick=delchg(<%=currentrow-1%>) > 
							<input name=op value="" type=hidden>
						</td>	      		
						<td align=center>
								<%if tmpRec(CurrentPage, x, 11)="" then %>
									<select name=groupid  >
									<option value="">---</option>
									<%sql="select * from  basiccode where func='groupid' and sys_type<>'AAA'  order by sys_type"
									set rds=conn.execute(sql)
									while not rds.eof 
									%>
									<option value="<%=rds("sys_type")%>" <%IF trim(tmpRec(CurrentPage,x,2))=rds("sys_type") THEN%> SELECTED<%END IF%>><%=rds("sys_value")%></option>
									<%rds.movenext
									wend
									set rds=nothing %> 
									</SELECT>	   
								<%else%>      			
									<%=trim(tmpRec(CurrentPage, x, 13))%>
									<input name=groupid type=hidden value="<%=trim(tmpRec(CurrentPage, x, 2))%>" >
								<%end if%>	   			
						</td>	
						<td align=center>
								<%if tmpRec(CurrentPage, x, 11)="" then %>
									<select name=zuno>
									<option value="">---</option>
									<%sql="select * from  basiccode where func='zuno' order by sys_type"
									set rds=conn.execute(sql)
									while not rds.eof 
									%>
									<option value="<%=rds("sys_type")%>" <%IF trim(tmpRec(CurrentPage,x,2))=rds("sys_type") THEN%> SELECTED<%END IF%>><%=rds("sys_value")%></option>
									<%rds.movenext
									wend
									set rds=nothing %> 
									</SELECT>	   
								<%else%>      			
									<%=trim(tmpRec(CurrentPage, x, 13))%>
									<input name=groupid type=hidden value="<%=trim(tmpRec(CurrentPage, x, 2))%>" >
								<%end if%>	   			
						</td>	
						<td nowrap>
								<input type="text" name=zcq value="<%=tmprec(1,x,5)%>" size=8 class=inputbox8  onchange="chkempid(<%=x-1%>)"  style='background-color:#ECF7FF;width:40%;height:22px' ondblclick="gotcq(<%=x-1%>)" >
								<input type="text" name=zcqname value="<%=tmprec(1,x,6)%>" size=17 class=readonly8 readonly    style='width:58%;height:22px'>
								<br><input type="text" name=zcqmail value="<%=tmprec(1,x,7)%>" size=28 class=inputbox8    style='width:100%;height:22px'>						
						</td>
						<td nowrap>
								<input type="text" name=jcq value="<%=tmprec(1,x,5)%>" size=8 class=inputbox8  onchange="chkempid(<%=x-1%>)"  style='background-color:#ECF7FF;width:40%;height:22px' ondblclick="gotcq(<%=x-1%>)" >
								<input type="text" name=jcqname value="<%=tmprec(1,x,6)%>" size=17 class=readonly8 readonly    style='width:58%;height:22px'>
								<br><input type="text" name=jcqmail value="<%=tmprec(1,x,7)%>" size=28 class=inputbox8    style='width:100%;height:22px'>						
						</td>
						<td nowrap>
							<input type="text" name=hcq value="<%=tmprec(1,x,5)%>" size=8 class=inputbox8  onchange="chkempid(<%=x-1%>)"  style='background-color:#ECF7FF;width:40%;height:22px' ondblclick="gotcq(<%=x-1%>)" >
							<input type="text" name=hcqname value="<%=tmprec(1,x,6)%>" size=17 class=readonly8 readonly    style='width:58%;height:22px'>
							<br><input type="text" name=hcqmail value="<%=tmprec(1,x,7)%>" size=28 class=inputbox8    style='width:100%;height:22px'>						
						</td>												
					</tr>
				  <%
				  next
				  %>	
				</table>							
				<input type=hidden value="" name=whsno>
				<input type=hidden value="" name=empid>  		
				<input type=hidden value="" name=empname>  
				<input type=hidden value="" name=levid>  
				<input type=hidden value="" name=unitno>
				<input type=hidden value="" name=country>  
				<input type=hidden value="" name=groupid>  
				<input type=hidden value="" name=job>  
				<input type=hidden value="" name=email1>  
				<input type=hidden value="" name=email2 >
			</td>
		</tr>
		<tr>
			<td align="center">
				<table class="table-borderless table-sm text-secondary txt">
					<tr><td align=center>Page:<%=currentpage%>/<%=totalpage%>, Count:<%=recordInDB%></td></tr>
				</table>
				<br>
				<table class="table-borderless table-sm text-secondary txt">
					<tr>
						<td align="left" width="50%">
					<% If CurrentPage > 1 Then %>
						<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
					<% Else %>
						<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
					<% End If %>
					
					<% If cint(CurrentPage) < cint(TotalPage) Then %>
						<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
					<% Else %>      
						<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">	
					<% End If %>
						</td>
						<td align="right" width="50%">
						<INPUT TYPE="button" name=send VALUE="CONFIRM" class="btn btn-sm btn-danger"  onClick="Go()"  >
						<INPUT TYPE="reset" name=send VALUE="CANCEL" class="btn btn-sm btn-outline-secondary"   >
						</TD>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
			

</FORM>

</BODY>
</HTML>
 

<script language=VBScript>
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function Q_Data()
   <%=self%>.TotalPage.value = ""
   <%=self%>.submit
end function

function gotcq(index)
	thiscols=document.activeElement.name
	open "getcqdata.asp?index="& index &"&cols1=" & thiscols ,"Back"
	parent.best.cols="65%,35%"
end function


function eidchg(index)
	eid = ucase(trim(<%=self%>.empid(index).value))
	if eid<>"" then 
		open "<%=self%>.back.asp?func=chkempid&code01="&eid&"&index="&index , "Back" 
		'parent.best.cols="70%,30%"
	end if 
end function

function emailchg(index)
	<%=self%>.op(index).value="upd"
end function 

function f()
	'<%=self%>.empid.focus()
end function 

function delchg(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="D"
	else
		<%=self%>.op(index).value=""
	end if 
end function  

FUNCTION go()	 	
	for zz= 1 to <%=self%>.pagerec.value 
		if <%=self%>.empid(zz-1).value<>"" then 
			if <%=self%>.whsno(zz-1).value="" then 
				alert "必須輸入資料!!"
				<%=self%>.whsno(zz-1).focus()
				exit function 
			end if 	
			if <%=self%>.levid(zz-1).value="" then 
				alert "必須輸入資料!!"
				<%=self%>.levid(zz-1).focus()
				exit function 
			end if 	 
			if <%=self%>.email1(zz-1).value="" and <%=self%>.email2(zz-1).value="" then 
				alert "至少須輸入ㄧ個Email!!"
				<%=self%>.email1(zz-1).focus()
				exit function 
			end if 							
		end if 
	next
	
	<%=self%>.action="<%=self%>.updateDB.asp"
	<%=self%>.submit()
end FUNCTION

function gotemp(index) 
	nfs = "levid"  
	open "../getempdata.asp?index="&index&"&formName="&"<%=self%>" &"&nfs="&nfs , "Back"
	parent.best.cols="50%,50%"
end function 

  

</script>


