<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%response.buffer=true%>
<%
Set conn = GetSQLServerConnection()
self="yeee04" 

if  instr(conn,"168")>0 then 
	w1="LA"
	w2 = "越南"	
elseif  instr(conn,"169")>0 then 
	w1="DN"	
	w2 = "同奈"	
elseif  instr(conn,"47")>0 then 
	w1="BC"	
	w2 = "越南"	
end if 	  

sortby = request("sortby") 
if sortby=""  then 
	sortby="B"
	sort_str = "a.empid, a.s_dat "	
elseif sortby="A" then 
	sort_str = "a.s_dat , a.empid"
elseif sortby="B" then 
	sort_str = "a.empid, a.s_dat "	
elseif sortby="C" then 	
	sort_str = "a.jb, a.empid, a.s_dat "	
end if 	

dd1=request("dd1")
dd2=request("dd2")
sothe=ucase(triM(request("sothe")))
ct = request("ct")
otherQ = trim(request("otherQ")) 
sts2 = trim(request("sts2"))   
if sts2="" then sts2="N"
whsno = trim(request("whsno"))  

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 20    'number of fields per record

	sql="select isnull(g.f_xid, a.xid) xid , isnull(g.empid, a.empid) empid , isnull(g.f_jb,a.jb) jb,  "&_
			"isnull(g.min_dat,a.s_dat)  s_dat, "&_
			"isnull(g.max_dat,a.e_dat)  e_dat,  "&_
			"case when isnull(g.empid,'')='' then a.s_tim else '08:00' end as s_tim ,"&_
			"case when isnull(g.empid,'')='' then a.e_tim else '17:00' end as e_tim ,"&_
			"case when isnull(g.empid,'')='' then a.hhour  else '8'  end as hhour ,"&_
			"case when isnull(g.empid,'')='' then a.jiadays  else g.dd  end as jiadays ,"&_
			"case when isnull(g.empid,'')='' then isnull(a.place,'')  else  isnull(g.f_place,'') end as nplace  ,"&_
			"c.empid as eid, c.empnam_cn, c.empnam_vn, groupid, gstr, job, jstr, d.sys_value as jbname ,"&_
			"mindat, maxdat , case when isnull(g.empid,'')='' then '' else 'S' end as sts, "&_
			"a.jb as old_jb , isnull(g.xjsts,'') xjsts  from "&_
			"( select * from [vyfynet].dbo.emptja where isnull(status,'')<>'D' )  a  "&_
			"join ( select   distinct xid, isnull(sts,'') sts, cqid, cqtype   from [vyfynet].dbo.empjiahj where cqtype='H' and isnull(sts,'')='Y' ) b on b.xid = a.xid   "&_
			"left join (select * from view_empfile ) c on c.empid = a.empid  "&_						
			"left join ( select  min(isnull(place,'')) f_place,  isnull(xjsts,'') xjsts, empid , jiatype as f_jb , xid as f_xid , min(convert(char(10),dateup,111)) as min_dat , "&_
			"max(convert(char(10),dateup,111)) as max_dat  , 	datediff(d , min(convert(char(10),dateup,111)) , max(convert(char(10),dateup,111))  )+1 dd "&_
			"from empHoliday where isnull(xid,'')<>'' group by empid, jiatype, xid ,  isnull(xjsts,'') ) g on isnull(g.f_xid,'') = a.xid   "&_			
			"left join ( select xid , min(s_dat) minDat , max(e_dat) Maxdat from [vyfynet].dbo.emptja where isnull(status,'')<>'D' group by xid , jb  ) f on f.xid = a.xid   "&_			
			"left join ( select * from  basicCode where func='JB'  ) d on d.sys_type = isnull(g.f_jb,a.jb)   "&_	
			"where c.country like '"& ct &"%' and a.empid like '"& sothe&"%' and   a.xid like '"&otherQ&"%' and c.whsno like '"&whsno&"%' " 			
	if dd1<>"" and dd2<>"" then 
		sql=sql&"and ( a.s_dat between '"& dd1&"' and '"&dd2&"' or a.e_dat between '"&dd1&"' and '"& dd2&"' ) "
	end if 			
	if sts2="*" then 
		sql=sql&"and  isnull(g.xjsts,'')='*'  "
	elseif sts2="N" then 
	sql=sql&"and  isnull(g.xjsts,'')<>'*'  "
	end if 
		sql=sql&"order by " & sort_str 
 '"where a.e_dat <= convert(char(10),getdate(),111 ) "&_ 
response.write sql
response.end
Set rs = Server.CreateObject("ADODB.Recordset")

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("xid"))		
				tmpRec(i, j, 2) = trim(rs("empid"))		
				tmpRec(i, j, 3) = trim(rs("empnam_cn"))		
				tmpRec(i, j, 4) = trim(rs("empnam_vn"))		
				tmpRec(i, j, 5) = trim(rs("jb"))		
				tmpRec(i, j, 6) = trim(rs("s_dat"))		
				tmpRec(i, j, 7) = trim(rs("s_tim"))						
				tmpRec(i, j, 8) = trim(rs("e_dat"))		
				tmpRec(i, j, 9) = trim(rs("e_tim"))		
				tmpRec(i, j, 10) = trim(rs("hhour"))		
				tmpRec(i, j, 11) = trim(rs("jiadays"))		
				tmpRec(i, j, 12) = trim(rs("jbname"))		
				tmpRec(i, j, 13) = trim(rs("nplace"))		
				tmpRec(i, j, 14) = trim(rs("minDat"))		
				tmpRec(i, j, 15) = trim(rs("Maxdat"))		
				tmpRec(i, j, 16) = trim(rs("sts"))		
				tmpRec(i, j, 17) = trim(rs("old_jb"))	
				tmpRec(i, j, 18) = trim(rs("xjsts"))	
				
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
	Session("YEBE0104B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBE0104B")

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

nowdate = year(date())&"/"&right("00"&month(date()),2)&"/"&right("00"&day(date()),2)

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2) 
END FUNCTION 
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
 
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=session("pgname")%></TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>

<table width=500  ><tr><td >
	<table width=450 BORDER=0 cellspacing="1" cellpadding="1" class=txt bgcolor=black>
	<tr bgcolor=#ffffff height=35>
		<Td align=center bgcolor="#ffcccc" width=150 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffcccc'"   style="cursor:hand" ><a href="yeee04.asp" target="_self" >銷假作業</a></td>
		<Td align=center bgcolor="#ffffff"  width=150 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yeee0401.asp" target="_self" >差旅費設定</a></td>
		<Td align=center bgcolor="#ffffff"  width=150 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"></td>		
	</tr>	
	</table>	
	
	<table  BORDER=0 cellspacing="1" cellpadding="2" class=txt8 width=640 > 	
		<tr>
			<Td align="right">日期<br>Ngay</td>
			<td><input name="dd1" size=11 class="inputbox8" value="<%=dd1%>" onblur="ddDatechg(1)">~<input name="dd2" size=11 class="inputbox8" value="<%=dd2%>" onblur="ddDatechg(2)" ></td>
			<Td align="right">工號<br>So the</td>
			<td><input name="sothe" size=6 maxlength="5" class="inputbox8" value="<%=sothe%>" ></td>
			<Td align="right">國籍<br>Quoc tich</td>
			<td>
				<select name="ct" class="txt8" style="width:80" onchange="gos()">
					<option value="">---</option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type<>'VN' ORDER BY SYS_type desc  "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if ct=rst("sys_type") then%> selected<%end if%>><%=RST("SYS_TYPE")%>  - <%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
				</select>				
			</td>
			<Td align="right">其他<br>Khac</td>
			<td><input name="otherQ" size=10  class="inputbox8" value="<%=otherQ%>" ></td>	
			<Td><input type="button"  name="btn" value="(S)查詢" class="button" onclick="gos()" onkeydown="gos()"></td>	
		</tr>			
		<tr>
			<Td align="right" nowrap>排序<br>sap xep</td>
			<td colspan=3>
				<select name="sortby" class="txt8" onchange="gos()">
					<option value="A" <%if sortby="A" then %>selected<%end if%>>theo ngay nhap phep 依日期</option>
					<option value="B" <%if sortby="B" then %>selected<%end if%>>theo so the 依工號</option>
					<option value="C" <%if sortby="C" then %>selected<%end if%>>theo loai phep 依假別</option>
				</select>
			</td>
			<Td align="right" nowrap>廠別<br>Xuong</td>
			<td >
				<select name="whsno" class="txt8" onchange="gos()">
					<option value=""  >--ALL--</option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>"<%if whsno=RST("SYS_TYPE") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
				</select>
			</td>
			<Td align="right" nowrap>狀態<br>Status</td>
			<td >
				<select name="sts2" class="txt8" onchange="gos()">
					<option value="" <%if sts2="" then %>selected<%end if%>>--ALL--</option>
					<option value="*" <%if sts2="*" then %>selected<%end if%>>(1)已銷假</option>
					<option value="N" <%if sts2="N" then %>selected<%end if%>>(2)未銷假</option>
				</select>
			</td>
		</tr>
	</table>	
	<table  BORDER=0 cellspacing="1" cellpadding="2" class=txt8 bgcolor="#999999">
		<tr height=25 bgcolor="#e4e4e4">
			<Td width=30 nowrap align="center">STT</td>
			<Td width=30 nowrap align="center" >銷假</td>
			<Td width=150 nowrap align="center" >銷假日期(起~迄)</td>
			<Td width=150 nowrap align="center" >工號/姓名</td>
			<Td width=30 nowrap  align="center">列印<br>In</td>			
			<Td width=50 nowrap  align="center">編號</td>			
			<Td width=60 nowrap  align="center">假別</td>
			<Td width=200 nowrap  align="center">原請假日期(起~迄)</td>
			<Td width=40 nowrap  align="center">天數</td>
			<Td width=40 nowrap  align="center">時數</td>
			<Td width=30 nowrap  align="center">地點</td>			
		</tr>
		<%response.flush%>
		<%for x = 1 to pagerec 
			if ( tmpRec(CurrentPage, x, 5)<>tmpRec(CurrentPage, x, 17)) or (tmpRec(CurrentPage, x, 6)<>tmpRec(CurrentPage, x, 14))  or (tmpRec(CurrentPage, x, 8)<>tmpRec(CurrentPage, x, 15))    then 
				ft_clr1="Blue" 
			else
				ft_clr1="Black" 
			end if			
			if tmpRec(CurrentPage, x, 1) <>"" then 
				old_jdays = datediff("d",tmpRec(CurrentPage, x, 14) , tmpRec(CurrentPage, x, 15))+1
			
		%>
			<tr bgcolor="#ffffff">
				<td align="center"><%=((currentpage-1)*pagerec)+x%></td>
				<td align="center">					
					<%if trim(tmpRec(CurrentPage, x, 18))<>"*"   then %> <!--未銷假-->
						<%if  trim(tmpRec(CurrentPage, x, 8)) <= nowdate  then    %>	
							<input type=checkbox name=func onclick="funcchg(<%=x-1%>)" <%if tmpRec(CurrentPage, x, 0)="Y" then%>checked<%end if%>>
						<%else%>	
							<input type="hidden"  name=func>
						<%end if%>	
					<%else%>						
						<%=trim(tmpRec(CurrentPage, x, 18))%>
						<input type="hidden"  name=func>
					<%end if%>
					
					<input type="hidden"  name="op" value="<%=tmpRec(CurrentPage, x, 0)%>">
					<input type="hidden"  name="xid" value="<%=tmpRec(CurrentPage, x, 1)%>">
					<input type="hidden"  name="empid" value="<%=tmpRec(CurrentPage, x, 2)%>">
				</td>
				<td nowrap>					
						<input name="dat1" class="readonly8" readonly  size=11  value="<%=tmpRec(CurrentPage, x, 6)%>"  > ~ 
						<input name="dat2" class="readonly8" readonly size=11 value="<%=tmpRec(CurrentPage, x, 8)%>"   >					
				</td>
				<td nowrap>		 <!-- emp Name-->				 
						<a href="vbscript:showholiday(<%=x-1%>)">
							<font color="<%=ft_clr1%>" ><%=tmpRec(CurrentPage, x, 2)%><br>
							<%=tmpRec(CurrentPage, x, 3)%><%=tmpRec(CurrentPage, x, 4)%>
							</font>
						</a>
				</td>
				<td nowrap align="center"> <!--print--> 				
					<%if trim(tmpRec(CurrentPage, x, 18))="*"  and  tmpRec(CurrentPage, x, 5)="I" then %>
					<a href="vbscript:printPhieu(<%=x-1%>)">Print</a>
					<%end if%>
				</td>
				<td nowrap align="center"><%=tmpRec(CurrentPage, x, 1)%></td><!--xid--> 				
				<td nowrap>
					<%=tmpRec(CurrentPage, x, 5)%>-<%=left(tmpRec(CurrentPage, x, 12),4)%>
				</td><!--假別-->
				<td nowrap>(<%=tmpRec(CurrentPage, x, 5)%>) 
						<%if tmpRec(CurrentPage, x, 14)<>"" and tmpRec(CurrentPage, x, 15)<>"" then %>
							 <%=tmpRec(CurrentPage, x, 14)%> ~ <%=tmpRec(CurrentPage, x, 15)%> , 共 <%=old_jdays%> 天
						<%end if 	%>
						<input type="hidden" name="b_jb" class="inputbox8" size=11  value="<%=tmpRec(CurrentPage, x, 5)%>">  
						<input type="hidden" name="B_dat1" class="inputbox8" size=11  value="<%=tmpRec(CurrentPage, x, 6)%>">  
						<input type="hidden" name="B_tim1" class="inputbox8" size=11  value="<%=tmpRec(CurrentPage, x, 7)%>">  
						<input type="hidden" name="B_dat2" class="inputbox8" size=11 value="<%=tmpRec(CurrentPage, x, 8)%>">
						<input type="hidden" name="B_tim2" class="inputbox8" size=11 value="<%=tmpRec(CurrentPage, x, 9)%>">
						<input type="hidden" name="mindat" class="inputbox8" size=11 value="<%=tmpRec(CurrentPage, x, 14)%>">
						<input type="hidden" name="maxdat" class="inputbox8" size=11 value="<%=tmpRec(CurrentPage, x, 15)%>">
				</td><!--請假日期-->
				<td nowrap align="center"><%=tmpRec(CurrentPage, x, 11)%></td><!--天數-->
				<td nowrap align="center"><%=tmpRec(CurrentPage, x, 10)%></td><!--時數-->
				<td nowrap align="center"><%=tmpRec(CurrentPage, x, 13)%></td><!--地點-->
			</tr>
			<%else%>
				<input type="hidden"  name="func" value="">
				<input type="hidden"  name="op" value="">
				<input type="hidden"  name="xid" value="">
				<input type="hidden"  name="empid" value="">
				<input type="hidden" name="dat1" class="inputbox8" size=11  value=""> 
				<input type="hidden" name="dat2" class="inputbox8" size=11 value="">					
				<input type="hidden" name="b_jb" class="inputbox8" size=11  value="">  
				<input type="hidden" name="B_dat1" class="inputbox8" size=11  value="">  
				<input type="hidden" name="B_tim1" class="inputbox8" size=11  value="">  
				<input type="hidden" name="B_dat2" class="inputbox8" size=11 value="">
				<input type="hidden" name="B_tim2" class="inputbox8" size=11 value="">		
				<input type="hidden" name="mindat" class="inputbox8" size=11 value="">
				<input type="hidden" name="maxdat" class="inputbox8" size=11 value="">
			<%end if%>
			<%f_xid = tmpRec(CurrentPage, x, 1) %>
		<%next%>
	</table>
	<input type="hidden"  name="func" value="">
	<input type="hidden"  name="op" value="">
	<input type="hidden"  name="xid" value="">
	<input type="hidden"  name="empid" value="">
	<input type="hidden" name="dat1" class="inputbox8" size=11  value=""> 
	<input type="hidden" name="dat2" class="inputbox8" size=11 value="">					
	<input type="hidden" name="b_jb" class="inputbox8" size=11  value="">  
	<input type="hidden" name="B_dat1" class="inputbox8" size=11  value="">  
	<input type="hidden" name="B_tim1" class="inputbox8" size=11  value="">  
	<input type="hidden" name="B_dat2" class="inputbox8" size=11 value="">
	<input type="hidden" name="B_tim2" class="inputbox8" size=11 value="">
	<input type="hidden" name="mindat" class="inputbox8" size=11 value="">
	<input type="hidden" name="maxdat" class="inputbox8" size=11 value="">	
	
	<TABLE WIDTH=500 class=txt8>
			<tr ALIGN=center>
		    <td align="left" height=40  >
			<% If CurrentPage > 1 Then %>
				<input type="submit" name="send" value="FIRST" class=button>
				<input type="submit" name="send" value="BACK" class=button>
			<% Else %>
				<input type="submit" name="send" value="FIRST" disabled class=button>
				<input type="submit" name="send" value="BACK" disabled class=button>
			<% End If %>
			<% If cint(CurrentPage) < cint(TotalPage) Then %>
				<input type="submit" name="send" value="NEXT" class=button>
				<input type="submit" name="send" value="END" class=button>
			<% Else %>
				<input type="submit" name="send" value="NEXT" disabled class=button>
				<input type="submit" name="send" value="END" disabled class=button>
			<% End If %>&nbsp;
			PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
			</td>
			<td>
				<input type=button name=btn value="(Y)Confirm"  class=button onclick="go()">				
			</td>
			</TR>
	</TABLE>	
		
</td></tr></table>

</body>
</html>
<!-- #include file="../Include/func.inc" -->

<script language=vbs>
function f()
	'<%=self%>.yymm.focus()
end function 
 
function funcchg(index)
	p = <%=self%>.pagerec.value  
	xid_str = trim(<%=self%>.xid(index).value) 
	
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="Y"		 
		
		dat1=<%=self%>.dat1(index).value
		dat2=<%=self%>.dat2(index).value 		
		open "<%=self%>.back.asp?func=yes&dat1="& dat1 & "&dat2="& dat2 &"&index="& index  &"&currentpage="& <%=currentpage%>   , "Back" 
		
		for zz = index to P-1 
			if zz+1 <=p-1 then 
				if xid_str = trim(<%=self%>.xid(zz+1).value) then 					
					<%=self%>.op(zz+1).value="Y"
					<%=self%>.func(zz+1).checked=true  					
					dat1=<%=self%>.dat1(zz+1).value
					dat2=<%=self%>.dat2(zz+1).value 							
					open "<%=self%>.back.asp?func=yes&dat1="& dat1 & "&dat2="& dat2 &"&index="& zz+1 &"&currentpage="& <%=currentpage%>   , "Back" 
					 
				end if 	
			end if 	
		next 		
	else
		<%=self%>.op(index).value=""		
		dat1=<%=self%>.dat1(index).value
		dat2=<%=self%>.dat2(index).value 		
		open "<%=self%>.back.asp?func=no&dat1="& dat1 & "&dat2="& dat2 &"&index="& index  &"&currentpage="& <%=currentpage%>    , "Back"  
		
		for zz = index to P-1 
			if zz+1 <=p-1 then 
				if xid_str = trim(<%=self%>.xid(zz+1).value) then 
					<%=self%>.op(zz+1).value=""
					<%=self%>.func(zz+1).checked=false
					dat1=<%=self%>.dat1(zz+1).value
					dat2=<%=self%>.dat2(zz+1).value 		
					open "<%=self%>.back.asp?func=no&dat1="& dat1 & "&dat2="& dat2 &"&index="& zz+1 &"&currentpage="& <%=currentpage%>   , "Back" 
				end if 	
			end if 	
		next 		
	end if  
	 
	
	'parent.best.cols="50%,50%" 
end function  

function datachg(index)
	dat1=<%=self%>.dat1(index).value
	dat2=<%=self%>.dat2(index).value 		
	open "<%=self%>.back.asp?func=datachg&dat1="& dat1 & "&dat2="& dat2 &"&index="& index  &"&currentpage="& <%=currentpage%>    , "Back"  
end function 

function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function 

function showholiday(index)
	wt = (window.screen.width )*0.8
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.05
	lt = (window.screen.availHeight)*0.1	
	empid = <%=self%>.empid(index).value
	dat1 = <%=self%>.dat1(index).value
	dat2 = <%=self%>.dat2(index).value

	open "<%=self%>.showdata.asp?empid1="& empid &"&dat1="& dat1 &"&dat2="& dat2     , "balnkN"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"
end function 

function go()	
 	<%=self%>.action="<%=self%>.upd.asp"
 	<%=self%>.submit()
end function

function goexcel()
	if <%=self%>.yymm.value="" then 
		alert "請輸入[統計年度]!!"
		<%=self%>.yymm.focus()
		exit function 
	end if 	
	'open "<%=self%>.toexcel.asp" , "Back" 
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	'parent.best.cols="50%,50%"
end function  

'*******檢查日期*********************************************
FUNCTION date_change(index,a)

if a=1 then
	INcardat = Trim(<%=self%>.dat1(index).value)
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2(index).value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1(index).value=ANS
			datachg(index)
		elseif a=2 then
			Document.<%=self%>.dat2(index).value=ANS
			datachg(index)
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.dat1(index).value=""
			Document.<%=self%>.dat1(index).focus()
		elseif a=2 then
			Document.<%=self%>.dat2(index).value=""
			Document.<%=self%>.dat2(index).focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION  

FUNCTION ddDatechg(a)

if a=1 then
	INcardat = Trim(<%=self%>.dd1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.dd2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dd1.value=ANS
		elseif a=2 then
			Document.<%=self%>.dd2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.dd1.value=""
			Document.<%=self%>.dd1.focus()
		elseif a=2 then
			Document.<%=self%>.dd2.value=""
			Document.<%=self%>.dd2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION  

function gos()
	<%=self%>.TotalPage.value = "" 
	<%=self%>.action="<%=self%>.Fore.asp"
	<%=self%>.submit()
end function  

function printPhieu(index)
	xid = <%=self%>.xid(index).value 
	empid = <%=self%>.empid(index).value 
	
	wt = (window.screen.width )*0.8
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.05
	lt = (window.screen.availHeight)*0.1	 

	open "<%=self%>.getrpt.asp?empid1="& empid &"&xid="&  xid  , "balnkN"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"
end function 
</script> 
<%response.end%>