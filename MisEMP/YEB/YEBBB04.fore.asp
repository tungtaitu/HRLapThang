<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="yebbb04"  


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

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 25    'number of fields per record    

yymm=request("yymm")
WHSNO=request("WHSNO")
country=request("country")
groupid=request("groupid")
empid1=request("empid1")
rpno=request("F_rpno")
rp_type = request("rp_type")
sortby = request("sortby")
if sortby="" then sortby="a.mdtm desc, a.empid "
queryx=trim(request("queryx"))
queryx = replace(replace(replace(queryx,"'","＂"),"%","％"),"+","＋")

if whsno="" and country="" and groupid="" and empid1="" and queryx=""  then 
	sql="select  top 1 c.sys_value,b.outdate, b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
		"(select * from emplicense where isnull(status,'')<>'D' and empid='XXX' ) a "&_
		"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
		"left join ( select *from basicCode where func='CDT'  ) c on  c.sys_type = a.cardData "&_		
		"order by " & sortby 
else	
	sql="select c.sys_value,b.outdate, b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
		"(select * from emplicense where isnull(status,'')<>'D' ) a "&_
		"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
		"left join ( select *from basicCode where func='CDT'  ) c on  c.sys_type = a.cardData "&_
		"where a.cdwhsno like '"& WHSNO &"%' and  b.groupid like '"& groupid &"%' and b.country like '"&country&"%' "&_
		"and ( a.empid like '%"& queryx &"%' or licensename like '%"&queryx &"%' )"&_
		"order by " & sortby 
end if 		 
Set rs = Server.CreateObject("ADODB.Recordset") 
'response.write sql
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 1, 3 
	IF NOT RS.EOF THEN 	
		'PageRec = rs.RecordCount+10 
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount+1
		gTotalPage = TotalPage
	END IF 	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rs("empid")				
				tmpRec(i, j, 2) = rs("licenseName")
				tmpRec(i, j, 2) = replace(tmpRec(i, j, 2),"+","＋") 
				tmpRec(i, j, 3) = rs("licenseorg")
				tmpRec(i, j, 4) = rs("qty")
				tmpRec(i, j, 5) = rs("cardData")
				tmpRec(i, j, 6) = rs("due_date")
				tmpRec(i, j, 7) = rs("empnam_cn")&rs("empnam_vn")
				tmpRec(i, j, 8) = ""
				tmpRec(i, j, 9) = rs("nindat")
				tmpRec(i, j, 10) = rs("whsno")
				tmpRec(i, j, 11) = rs("groupid")
				tmpRec(i, j, 12) = rs("zuno")
				tmpRec(i, j, 13) = rs("gstr")
				tmpRec(i, j, 14) = rs("job")
				tmpRec(i, j, 15) = rs("jstr")
				tmpRec(i, j, 16) = rs("outdate")
				tmpRec(i, j, 17) = tmpRec(i, j, 5)&rs("sys_value")
				tmpRec(i, j, 18) = rs("licenseNo")				
				if trim(rs("outdate"))="" then 
					tmpRec(i, j, 19)="black"
				else
					tmpRec(i, j, 19)="red"
				end if 	
				tmpRec(i, j, 20)=rs("period_dat")
				tmpRec(i, j, 21)=rs("amt")					
				tmpRec(i, j, 22)=rs("cardmemo")			
				if trim(rs("outdate"))="" then 
					tmpRec(i, j, 23)=""
				else
					tmpRec(i, j, 23)=rs("outdate")&"Da thoi viec"
				end if 
				tmpRec(i, j, 24)=rs("autoid")
				rs.MoveNext 
			else 
				'exit for 
				tmpRec(i, j, 1)=""
				tmpRec(i, j, 2) = ""
				tmpRec(i, j, 4)="1"
			end if 
	 	next 
	
	 	if rs.EOF then 
			rs.Close 
			Set rs = nothing
			exit for 
	 	end if 
	next
	Session("yebbb04B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	StoreToSession()	  
	tmpRec = Session("yebbb04B")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(TotalPage) then
			     CurrentPage = CurrentPage + 1
			  else
			  	 CurrentPage = TotalPage
		      end if
	     Case "END"
		      CurrentPage = TotalPage
	     Case Else
		      CurrentPage = 1
	end Select	
end if	
	

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">   
</head>  
<body  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=hidden NAME=sortby VALUE="<%=sortby%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table  class="txt" cellpadding=3 cellspacing=3> 		 
					<TR>
						<TD align=right height=30 >廠別<BR><font class="txt8">Xuong</font> </TD>
						<TD> 
							<select name=WHSNO  onchange="go()" style="width:120px">
								<option value="">全部ALL</option>
								<%
								if session("rights")=0 then 
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("netwhsno") &"' ORDER BY SYS_TYPE "
								end if 	
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								rst.close					
								%>
							</SELECT>
							<%SET RST=NOTHING %>
						</TD>  	 
						<TD  align=right height=30 >國籍<BR><font class="txt8">Quoc Tich</font></TD>
						<TD >
							<select name=country    style="width:120px" onchange="go()" >
								<option value="">全部 </option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=country then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND
								rst.close					
								%>
							</SELECT>
							<%SET RST=NOTHING %>
							<%
							'conn.close
							'set conn=nothing
							%>
							
						</TD>					
						<TD nowrap align=right >部門<BR><font class="txt8">Don vi</font></TD>
						<TD >
							<select name=GROUPID    onchange="go()"  style="width:120px">
							<option value="">全部 </option>
							<%
							SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
							'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							'RESPONSE.WRITE SQL 
							WHILE NOT RST.EOF  
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
							<%
							RST.MOVENEXT
							WEND 
							%>
							</SELECT>
							<%SET RST=NOTHING %>
							<%
							conn.close
							set conn=nothing
							%>
						</td>
					</tr>								
					<tr>	
						<td align=right>關鍵字</td>
						<td colspan=3>
							<input type="text" name=queryx size=30  maxlength=255 value='<%=queryx%>'>			
						</td>
						<td colspan=2>
							<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="(S)查詢K.tra" onclick="go()" onkeydown="go()">	&nbsp;						
							<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="(rs)重新查詢K.tra" onclick="gon()"  >
						</td>
					</TR> 		
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table  id="myTableGrid" width="98%">
					<tr class="header">
						<Td width="15px">xoa</td>
						<Td width=15>STT</td>				
						<Td onclick="dchg(3)" style='cursor:hand' >工號<br>so the<br><img src="../picture/soryby.gif"></td>
						<Td onclick="dchg(4)" style='cursor:hand'>部門<br>bo phan<br><img src="../picture/soryby.gif"></td>										
						<Td >姓名<br>ho ten<br></td>		
						<Td >證照名稱<br>Hang muc</td>
						<Td >證號<br>ma so</td>
						<Td >發證機關<br>Cơ Quan Cấp Chứng Nhận</td>
						<Td >發證日期<br>ngay cap chung chi</td>		
						<Td >份數<br>S.L</td>
						<Td >版本<br>Loai</td>
						<Td >有效期<br>Co hieu luc(y/m/d)</td>
						<Td >費用<br>Phi(VND)</td>
						<Td >備註<br>Ghi Chu</td>		
					</tr>
					<%for x =  1 to  pagerec						 	
						if tmpRec(CurrentPage, x, 0)="upd" or tmpRec(CurrentPage, x, 1)=""  then %>		
							<Tr>
								<td>
									<input name=func type=hidden>
									<input name=op type=hidden value="<%=tmpRec(CurrentPage, x, 0)%>" >
								</td>
								<Td align=center><%=(currentpage-1)*pagerec+x%></td>
								<Td align=center>
									<input type="text" style="width:98%" name="empid"  value="<%=tmpRec(CurrentPage, x, 1)%>"  size=5 ondblclick='gotemp(<%=x-1%>)' onblur="empidchg(<%=x-1%>)" maxlength=5>
								</td>
								<Td align=center>
									<input type="text" style="width:98%" name="F_groupid"  value="<%=tmpRec(CurrentPage, x, 13)%>"  readonly class="readonly" size=6 >
								</td>
								<Td align=left>
									<input type="text" style="width:98%" name="empname"  value="<%=tmpRec(CurrentPage, x, 7)%>"  readonly class="readonly" size=20 >
								</td>
								<Td align=left>
									<input type="text" style="width:98%" name="cardName"  value="<%=tmpRec(CurrentPage, x, 2)%>"   size=28 maxlength=255  onblur="datachg(<%=x-1%>)">
								</td>				
								<Td align=left>
									<input type="text" style="width:98%" name="cardno"  value="<%=tmpRec(CurrentPage, x, 18)%>"    size=18 maxlength=255 onblur="datachg(<%=x-1%>)">
								</td>				
								<Td align=left>
									<input type="text" style="width:98%" name="cardorg"  value="<%=tmpRec(CurrentPage, x, 3)%>"    size=22 maxlength=255 onblur="datachg(<%=x-1%>)">
								</td>	
								<Td align=left>
									<input type="text" style="width:98%" name="carddat"  value="<%=tmpRec(CurrentPage, x, 6)%>"    size=10 onblur="datachg(<%=x-1%>)" >
								</td>	
								<Td align=left>
									<input type="text" style="width:98%" name="qty"  value="<%=tmpRec(CurrentPage, x, 4)%>"    size=2 onblur="datachg(<%=x-1%>)" >
								</td>	
								<td align=left>
									<select name="cardData"   style="width:80px"  onchange="datachg(<%=x-1%>)">
										<option value="">----</option>
										<option value="A" <%if tmpRec(CurrentPage, x, 5)="A" then%>selected<%end if%>>A正本</option>
										<option value="B" <%if tmpRec(CurrentPage, x, 5)="B" then%>selected<%end if%>>B影本(以公証)</option>
										<option value="C" <%if tmpRec(CurrentPage, x, 5)="C" then%>selected<%end if%>>C影本</option>
									</select>
								</td>	
								<Td align=left>
									<input type="text" style="width:98%" name="period_date"  value="<%=tmpRec(CurrentPage, x, 20)%>"  size=10 onblur="datachg(<%=x-1%>)" >
								</td>	
								<Td align=left>
									<input type="text" style="width:98%" name="amt"  value="<%=tmpRec(CurrentPage, x, 21)%>"    size=10 onblur="datachg(<%=x-1%>)" >
								</td>	
								<Td align=left>
									<input type="text" style="width:98%" name="cardmemo"  value="<%=tmpRec(CurrentPage, x, 22)%>"    size=15 onblur="datachg(<%=x-1%>)" >
								</td>
							</tr>		
						<%else%>
							<tr>
								<td>
									<%if tmpRec(CurrentPage, x, 0)="del" then %>
										<input name=func type=checkbox onclick="del(<%=x-1%>)" checked >
										<input name=op type=hidden value="del">
									<%else%>
										<input name=func type=checkbox onclick="del(<%=x-1%>)">
										<input name=op type=hidden value="no">
									<%end if%>	
									<input name=empid type=hidden value="<%=tmpRec(CurrentPage, x, 1)%>">
									<input name=F_groupid type=hidden value="<%=tmpRec(CurrentPage, x, 13)%>">
									<input name=empname type=hidden value="<%=tmpRec(CurrentPage, x, 7)%>">
									<input name=cardName type=hidden value="<%=tmpRec(CurrentPage, x, 2)%>">
									<input name=cardno type=hidden value="<%=tmpRec(CurrentPage, x, 18)%>">
									<input name=cardorg type=hidden value="<%=tmpRec(CurrentPage, x, 3)%>">
									<input name=carddat type=hidden value="<%=tmpRec(CurrentPage, x, 6)%>">
									<input name=qty type=hidden value="<%=tmpRec(CurrentPage, x, 4)%>">
									<input name=cardData type=hidden value="<%=tmpRec(CurrentPage, x, 5)%>">
									<input name=period_date type=hidden value="<%=tmpRec(CurrentPage, x, 20)%>">
									<input name=amt type=hidden value="<%=tmpRec(CurrentPage, x, 21)%>">
									<input name=cardmemo type=hidden value="<%=tmpRec(CurrentPage, x, 22)%>">
								</td>
								<Td align=center><%=(currentpage-1)*pagerec+x%></td>
								<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 1)%></font></td>
								<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 13)%></font></td>				
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 7)%></font></td><!--姓名-->			
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 2)%></font></td>
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 18)%></font></td>
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 3)%></font></td>
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 6)%></font></td>
								<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 4)%></font></td>
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 17)%></font></td>
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 20)%></font></td>
								<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 21)%></font></td>
								<Td align=left>
									<font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 22)%>
									<%if trim(tmpRec(CurrentPage, x, 16))<>""then%>
										<%=tmpRec(CurrentPage, x, 16)&" 離職(Thoi viec)"%>
									<%end if%>	
									</font>
								</td>
							</tr>
						<%end if%>
					<%next%>
				</table>
			</td>
		</tr>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td align="center">
				<input type=hidden name=empid>
				<input type=hidden name=rpno>
				<table class="table-borderless table-sm bg-white text-secondary">
					<Tr>
					<td align="CENTER" height=40  >    
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
					<Td height=25 align=center>page : <%=currentpage%>/<%=totalpage%>, recordCount:<%=recordIndB%></td>
					<td>
						<%if session("mode")="W" then %>
							<input type=button name="btn" value="(Y)Confirm"  class="btn btn-sm btn-danger" onclick=goupd() >
							<input type=button name="btn" value="(N)Cnacel"  class="btn btn-sm btn-outline-secondary" onclick="goclr()">
							<input type=button name="btn" value="轉Excel"  class="btn btn-sm btn-outline-secondary" onclick="goexcel()">
						<%end if%>	
					</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
			

</form>
</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("yebbb04B")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 1) = request("empid")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("cardName")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 3) = request("cardorg")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 4) = request("qty")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 5) = request("cardData")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 6) = request("carddat")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 7) = request("empname")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 13) = request("F_groupid")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 18) = request("cardno")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 20) = request("period_date")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 21) = request("amt")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 22) = request("cardmemo")(CurrentRow) 
	next 
	Session("yebbb04B") = tmpRec
End Sub
%>
<script language=vbs>  
function empidchg(index)
	if trim(<%=self%>.empid(index).value)<>"" then
		eidstr = Ucase(trim(<%=self%>.empid(index).value))
		open "<%=self%>.back.asp?func=chkemp&code1="& eidstr & "&index="& index &"&currentpage="& <%=currentpage%>, "Back" 
		'parent.best.cols="70%,30%" 
		'datachg(index)
	end if 
end function 
function del(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="del"
		open "<%=self%>.back.asp?func=del&index="& index &"&currentpage="& <%=currentpage%>, "Back" 
	else
		<%=self%>.op(index).value=""
		open "<%=self%>.back.asp?func=no&index="& index &"&currentpage="& <%=currentpage%>, "Back" 
	end if 
end function 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function 

function datachg(index)
	<%=self%>.op(index).value="upd"
	c1=escape(Ucase(Trim(<%=self%>.empid(index).value)))
	c2=escape(Ucase(Trim(<%=self%>.cardName(index).value)))	
	c3=escape(Ucase(Trim(<%=self%>.cardorg(index).value)))	
	c4=escape(Ucase(Trim(<%=self%>.qty(index).value)))
	c5=escape(Ucase(Trim(<%=self%>.cardData(index).value)))
	c6=escape(Ucase(Trim(<%=self%>.carddat(index).value)))
	c7=escape(Ucase(Trim(<%=self%>.cardno(index).value)))
	c8=escape(Ucase(Trim(<%=self%>.period_date(index).value)))	
	c9=escape(Ucase(Trim(<%=self%>.amt(index).value)))
	c10=escape(Ucase(Trim(<%=self%>.cardmemo(index).value)))
	
	open "<%=self%>.back.asp?func=upd&currentpage="& <%=currentpage%> &_
		 "&index="& index &_
		 "&c1="& c1 &_
		 "&c2="& c2 &_
		 "&c3="& c3 &_
		 "&c4="& c4 &_
		 "&c5="& c5 &_
		 "&c6="& c6 &_
		 "&c7="& c7 &_
		 "&c8="& c8 &_
		 "&c9="& c9 &_
		 "&c10="& c10 ,"Back"	
	'parent.best.cols="70%,30%"	 	 
end function 

function goclr()	
	
	<%=self%>.totalpage.value="0"
	<%=self%>.submit()
end function  

function gon()
	open "<%=self%>.Fore.asp", "_self"
end function

function gotemp(index)
	'alert index 
	nfs="cardName"  'next focus
	open "../getempdata.asp?formName="&"<%=self%>"&"&index=" & index &"&nfs="& nfs , "Back"
	parent.best.cols="60%,40%"
end function 

function dchg(a) 
	select case a 
		case 1 
			<%=self%>.sortby.value="rpno desc"
		case 2 
			<%=self%>.sortby.value="rp_dat, rpno "
		case 3 
			<%=self%>.sortby.value="a.empid"
		case 4 
			<%=self%>.sortby.value="b.groupid, a.empid"
		case 5 
			<%=self%>.sortby.value="rp_func, rpno "
	end select 	
	<%=self%>.totalpage.value="0"
 	<%=self%>.action="<%=self%>.Fore.asp"
 	<%=self%>.submit() 														
end function  

function f()
	'<%=self%>.yymm.focus()	
	'<%=self%>.country.SELECT()
end function

function gonew()    
	open "<%=self%>.new.asp", "_self"
end function  

function go2(a) 
	if a=1 then 
		if <%=self%>.yymm.value<>"" then 
			<%=self%>.totalpage.value="0"
		 	<%=self%>.action="<%=self%>.Fore.asp"
		 	<%=self%>.submit() 
		 end if 	
	elseif a=2 then 
		if <%=self%>.empid1.value<>"" then 
			<%=self%>.totalpage.value="0"
		 	<%=self%>.action="<%=self%>.Fore.asp"
		 	<%=self%>.submit() 
		 end if 	
	end if  
end function  
	
function go() 	
	<%=self%>.totalpage.value="0"
 	<%=self%>.action="<%=self%>.Fore.asp"
 	<%=self%>.submit() 
end function  

function goupd() 		
 	<%=self%>.action="<%=self%>.upd.asp"
 	<%=self%>.submit() 
end function 

function chkdata(index)
	 rpno= <%=self%>.rpno(index).value
	 empid = <%=self%>.empid(index).value
	 open "<%=self%>.foregnd.asp?rpno="&rpno &"&empid="&empid, "_self" 
	 
end function   
	

'*******檢查日期*********************************************
FUNCTION date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS			
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS		 			
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION
</script> 