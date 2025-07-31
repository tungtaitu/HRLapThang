<%@language=vbscript CODEPAGE=65001%>
<!---------  #include file="../GetSQLServerConnection.fun"  -------->
<!--#include file="../include/sideinfo.inc"-->
<%
Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd
'session.codepage=65001
SELF = "yecb0201"
Set conn = GetSQLServerConnection()

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 99    'number of fields per record
f_w1 = request("f_w1")
f_ct = request("f_ct")
'on error resume next

'sqln="select * from scode_big where  tblid='"& DB_TBLID &"' "
'set rds=conn.execute(sqln)
'scodeBig = rds("description") f_w1



sql="select   max(isnull(yymm,'')) yymm, max(isnull(dm,'')) dm, code1=max(substring(code,2,1))  ,  code2=max(right(code,1))  ,  "&_
		"Flines=CASE WHEN ASCII(max(substring(code,2,1)))>=65 then ASCII(max(substring(code,2,1)))-55 else max(substring(code,2,1)) end  , "&_
		"Frows=CASE WHEN ASCII(max(right(code,1)))>=65 then ASCII(max(right(code,1)))-55 else max(right(code,1))end "&_
		"from empsalarybasic where func='AA'    and country='vn' "&_
		"and bwhsno='"&f_w1&"'  and bonus>0  "
set rds=conn.execute(Sql)  
'response.write sql 
'response.end 
Flines = rds("Flines")
Frows = rds("Frows")
endym=rds("yymm")
dm=rds("dm")
'response.write "<p>"&Frows &"<BR>"&Flines &"<BR>"
'response.end 
 		
if request("TotalPage") = "" then
   CurrentPage = 1  

 	source = "exec proc_sp_ycb02 '"&f_w1&"' " 
	'response.write  source &"<BR>"
	'response.end
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open Source, conn, 3, 3
	if not rs.eof then
		while not rs.eof 
			kk=kk+1 
			rs.movenext
		wend 	
		flines= kk+5	
		PageRec = flines
		'rs.PageSize = PageRec
		RecordInDB = flines
		TableRec = Frows+9
		TotalPage =1 
		gTotalPage = totalpage 
		 
	end if
	rs.movefirst
	'response.end

	'Set conn = nothing
	'response.write gTotalPage & pagerec & TableRec &"<BR>"

	Redim tmpRec(gTotalPage, PageRec, TableRec+5)   'Array
	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then 
				tmpRec(i, j, 0) = "no"
				for ix = 1 to Frows+4 			
					tmpRec(i, j, ix) = rs(ix-1) 						
					'response.write "ix="& ix &"-"& tmpRec(i, j, ix)&"<BR>" 					 
				next 					
 			else			
				exit for
			end if 		 
			
			rs.movenext
		next
		
		'response.write tmpRec(i, j, k)&"<BR>" 	

		if rs.EOF then
			rs.Close
			Set rs = nothing
			exit for
		end if
	next
	  
 
end if
%>

<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

function m(index)
   <%=SELF%>.send(index).style.backgroundcolor="lightyellow"
   <%=SELF%>.send(index).style.color="red"
end function

function n(index)
   <%=SELF%>.send(index).style.backgroundcolor="khaki"
   <%=SELF%>.send(index).style.color="black"
end function 

'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

</SCRIPT>
</head>
<body   topmargin=40    onkeydown="enterto()" >  
<form method="post" name="<%=SELF%>" >
 
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=TableRec VALUE="<%=TableRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<input TYPE=hidden name=f_w1 value="<%=request("f_w1")%>">
<input TYPE=hidden name=f_ct value="<%=request("f_ct")%>">
<input TYPE=hidden name=flines value="<%=Flines%>">
<input TYPE=hidden name=frows value="<%=frows%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td align="right">(薪資專案)截止年月<br>End (yymm)</td>
									<td><input type="text" style="width:100px" class="txt8" name="endym" maxlength=6 value="<%=endym%>"></td>
									<td  align="right">幣別<br>loai Tien</td>
									<td><input type="text" style="width:100px" class="txt8" name="dm"  maxlength=3 value="<%=dm%>">(VND or USD)</td>		
									<td  align="right">廠別<br>loai xuong</td>
									<td bgcolor="#e4e4e4" width=40 nowrap align="Center"><%=F_w1%></td>		
									<td  align="right">國籍<br>Quoc tich</td>
									<td bgcolor="#e4e4e4" width=40 nowrap align="Center"><%=F_ct%></td>	
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%">
							  <tr class="header" height="35px">
								<td width=30 align=center >Code</td>        
								<td  align=center  >說明</td> 
										<td  align=center  >B0</td> 
										<%for x= 1 to frows+5
											'if x >=10 then 	showstr=chr(x+55) else showstr=x 					
										%>
											<td  align=center  >B<%=x%></td> 
										<%next%>
										
							  </tr>
							  <%
								for CurrentRow = 1 to PageRec

								j = 1
								if j=1 then
									wk_color = "#E1E8F0"
									j = 0
								else
									wk_color = "#EBEDF1"
									j = 1
								end if 
								if tmprec(currentpage,CurrentRow,1)="" and tmprec(currentpage,CurrentRow,2)="" then
									code_lins="Z"&CurrentRow
								else
									code_lins=tmprec(currentpage,CurrentRow,1)&tmprec(currentpage,CurrentRow,2)
								end if	
							  %>
							  <TR BGCOLOR="#FFFFFF">
								 <td> 
										 <input type="text" style="width:98%" name="line_code"  value="<%=code_lins%>">
										 </td>
										 <td>					
											  <input type="text" style="width:98%" name="descp"  value="<%=tmprec(currentpage,CurrentRow,3)%>">
											</td>
										 <td><input type="text" style="width:98%" name="B0" class="txt8" value="<%=formatnumber(tmprec(currentpage,CurrentRow,4),0)%>"> 
										 
										 </td>
										<%for y= 5 to frows+4%> 
											<td>
												<input type="text" style="width:98%" name="Bx" class="txt8" value="<%=formatnumber(tmprec(currentpage,CurrentRow,y),0)%>" onblur="chkbsry(<%=((CurrentRow-1)*(TableRec-4))+(y-5)%>)"  > 						
											</td> 
										<%next%>
										<%for z=  frows+5 to TableRec%> 
											<td> 
												<input type="text" style="width:98%" name="Bx" class="txt8" value="<%=formatnumber(tmprec(currentpage,CurrentRow,z),0)%>" onblur="chkbsry(<%=((CurrentRow-1)*(TableRec-4))+(z-5)%>)"> 						
											</td> 
										<%next%>
										<%
										allnum=0
										if tmprec(currentpage,CurrentRow,4)="" then allnum=allnum+0 else allnum=allnum+cdbl(tmprec(currentpage,CurrentRow,4))
										for yy = 5 to frows+4
											if tmprec(currentpage,CurrentRow,yy)="" then  
												allnum=allnum+0 
											else
												allnum=allnum+tmprec(currentpage,CurrentRow,yy) 
											end if
										next
										for zz = frows+5 to TableRec 
											if tmprec(currentpage,CurrentRow,zz)="" then  
												allnum=allnum+0 
											else
												allnum=allnum+tmprec(currentpage,CurrentRow,zz) 
											end if
										next
										'response.write allnum
										tmprec(currentpage,CurrentRow,TableRec+1)=allnum
										%> 				
										<input name="chkNum" type="hidden" value="<%=allnum%>"> 				
							  </TR>
							  <%next%>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>	 
								  <td align="CENTER"    height=30 >
									<%if UCASE(session("mode"))="W" then%>
										<input type="button" name="send" value="(Y)Confirm" onclick="GO()" class="btn btn-sm btn-danger">
										<input type="reset" name="send" value="(N)Cancel" class="btn btn-sm btn-outline-secondary" >
									<%end if%>	      
								  </td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
 
</form>

</body>
</html>
 

<script language=vbscript>  

function Go()
   <%=SELF%>.action = "<%=SELF%>.upd.asp"
   <%=SELF%>.submit
end function

function Clear()
	open "<%=SELF%>.asp", "_self"
end function


function chkbsry(index)
	if trim(<%=self%>.bx(index).value)<>"" then 
		if  isnumeric(<%=self%>.bx(index).value)=false then 
			alert "xin danh lai so 請輸入數字"
			<%=self%>.bx(index).value="0"
			<%=self%>.bx(index).select()
			exit function 
		elseif 	cdbl(<%=self%>.bx(index).value)<0 then 
			alert "xin danh lai so 請輸入數字,必須>0"
			<%=self%>.bx(index).value="0"
			<%=self%>.bx(index).select()
		else	
			<%=self%>.bx(index).value=formatnumber(<%=self%>.bx(index).value,0)
		end if 
	end if 
end function 
 
</script>


