<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->

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

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 20    'number of fields per record    

yymm=request("yymm")
yymm2=request("yymm2")
if request("yymm2")="" then yymm2=yymm
whsno=request("whsno")
country=request("country")
groupid=request("groupid")
empid1=request("empid1")
rpno=request("F_rpno")
rp_type = request("rp_type")
sortby = request("sortby")
if sortby="" then sortby="rpno desc"

if yymm="" and whsno="" and country="" and groupid="" and empid1="" and rpno="" then 
	sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
		"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112)='xxx' ) a "&_
		"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
		"left join ( select *from basicCode ) c on c.func=case when a.rp_type='R' then 'goods' else case when a.rp_type='P' then 'bads' else '' end end  and c.sys_type = a.rp_func "&_
		"where a.rpwhsno like '"& session("rpwhsno") &"%' and  b.groupid like '"& groupid &"%' and b.country like '"&country&"%' "&_
		"and a.empid like '"& empid1 &"%' and a.rp_type like '"&rp_type&"%' "&_
		"order by " & sortby 
else 
	if yymm<>""  then 
		sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
				"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112) between '"& yymm &"' and '"&yymm2&"' ) a "
	else 
		sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
				"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112) like  '"& yymm &"%'  ) a  "
	end if  

	sql=sql&"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
			"left join ( select *from basicCode ) c on c.func=case when a.rp_type='R' then 'goods' else case when a.rp_type='P' then 'bads' else '' end end  and c.sys_type = a.rp_func "&_
			"where a.rpwhsno like '"& WHSNO &"%' and  b.groupid like '"& groupid &"%' and b.country like '"&country&"%' "&_
			"and a.empid like '%"& empid1 &"%' and left(rpno,2) like '"& left(rpno,2) &"%'and a.rp_type like '"&rp_type&"%' "&_
			"order by  " & sortby 	
end if 
Set rs = Server.CreateObject("ADODB.Recordset") 
'response.write sql
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 1, 3 
	IF NOT RS.EOF THEN 	
		'PageRec = rs.RecordCount 
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
				tmpRec(i, j, 1) = rs("rpno")
				tmpRec(i, j, 2) = rs("rp_date")
				tmpRec(i, j, 3) = rs("empid")
				tmpRec(i, j, 4) = rs("rp_type")
				tmpRec(i, j, 5) = rs("rp_func")
				tmpRec(i, j, 6) = rs("rp_method")
				tmpRec(i, j, 7) = rs("empnam_cn")
				tmpRec(i, j, 8) = rs("empnam_vn")
				tmpRec(i, j, 9) = rs("nindat")
				tmpRec(i, j, 10) = rs("whsno")
				tmpRec(i, j, 11) = rs("groupid")
				tmpRec(i, j, 12) = rs("zuno")
				tmpRec(i, j, 13) = rs("gstr")
				tmpRec(i, j, 14) = rs("job")
				tmpRec(i, j, 15) = rs("jstr")
				if rs("rp_type")="R" then
					tmpRec(i, j, 16)=rs("rp_type")&" 獎勵"
				elseif rs("rp_type")="P" then 
					tmpRec(i, j, 16)=rs("rp_type")&" 懲罰"
				else
					tmpRec(i, j, 16)=""
				end if 	
				tmpRec(i, j, 17) = tmpRec(i, j, 5) & " " & rs("sys_value")
				tmpRec(i, j, 18)=RS("FILENO")
				if rs("rp_type")="P" then 
					tmpRec(i, j, 19)="black"
				else
					tmpRec(i, j, 19)="blue"
				end if 	
				tmpRec(i, j, 20) = left(rs("rpmemo"),15)&"..."
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
	Session("yebbb03B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("yebbb03B")

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
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=hidden NAME=sortby VALUE="<%=sortby%>">

<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=session("pgname")%></TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>		
<table width=550  ><tr><td >
	<table  align=center border=0  cellspacing="1" cellpadding="3" class=txt8 > 		 
		<TR>
			<TD nowrap align=right height=30 >廠別<BR><font class=txt8>Xuong</font> </TD>
			<TD > 
				<select name=WHSNO  class=txt8  >
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
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=session("mywhsno") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>  			 
		 <TD nowrap align=right height=30 >統計<br>年月</TD>
			<TD colspan=3>
				<input name=yymm class=inputbox size=8  value="<%=yymm%>"  maxlength=6 >~
				<input name=yymm2 class=inputbox size=8  value="<%=yymm2%>" maxlength=6 > 
				(yyyymm)
			</TD>	  

			<!--TD nowrap align=right height=30 >獎懲<BR><font class=txt8>&nbsp;</font></TD>
			<td  >
				<select name="rp_type" class=txt8 onchange="go()" style='width:70'>
					<option value="">---</option>
					<option value="R" <%if rp_type="R" then%>selected<%end if%>>R 獎勵(OK)</option>
					<option value="P" <%if rp_type="P" then%>selected<%end if%>>P 懲罰(NG)</option>
				</select>
			</td-->				
			<td>	
				<input type=button name=btn value="(N)資料新增" class=button style='background-color:#ffccff' onclick=gonew() >
			</td>
		</tr>
		<tr>
		 	<TD nowrap align=right height=30 >國籍<BR><font class=txt8>Quoc Tich</font></TD>
			<TD >
				<select name=country  class=txt8 style='width:85'   >
					<option value="">----</option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=country then%>selected<%end if%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>					
			<TD nowrap align=right >部門<BR><font class=txt8>Don vi</font></TD>
			<TD >
				<select name=GROUPID  class=txt8  >
				<option value="" selected >----</option>
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
				<%rst.close
				SET RST=NOTHING 
				conn.close 
				set conn=nothing
				%>
			<td nowrap align=right >工號<BR><font class=txt8>So The</font></td>
			<td  >
				<input name=empid1 class=inputbox size=10 maxlength=5   value="<%=empid1%>"></td>
			<td>	
				<input type=button  name=btm class=button value="(S)查詢K.tra" onclick="go()" onkeydown="go()">
			</td>
		</TR>  
	</table> 
</td></tr></table> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>		
<table BORDER=0 cellspacing="1" cellpadding="2"  class=txt8 bgcolor=#cccccc >
	<tr height=25 bgcolor=#e4e4e4  >
		<Td width=30 nowrap align=center >STT</td>
		<Td width=40 nowrap align=center >獎懲<br>thuong phat</td>
		<Td width=50 nowrap align=center onclick="dchg(1)" style='cursor:hand'>編號<br>so<br><img src="../picture/soryby.gif"></td>
		<Td width=60 nowrap align=center onclick="dchg(2)" style='cursor:hand'>事件日期<br>Ngay<br><img src="../picture/soryby.gif"></td>
		<Td width=50 nowrap align=center onclick="dchg(3)" style='cursor:hand' >工號<br>so the<br><img src="../picture/soryby.gif"></td>
		<Td width=120 nowrap align=center >姓名<br>ho ten<br></td>
		<!--Td width=70 nowrap align=center >到職日<br>NVX<br></td-->
		
		<Td width=50 nowrap align=center onclick="dchg(4)" style='cursor:hand'>部門<br>bo phan<br><img src="../picture/soryby.gif"></td>		
		<Td width=70 nowrap align=center onclick="dchg(5)" style='cursor:hand'>方式<br>phuong thuc<br><img src="../picture/soryby.gif"></td>
		<Td width=120 nowrap align=center  >內容<br>thuyet minh</td>
		<Td width=120  align=center  >處理說明<br>phuong thuc xu ly </td>		
		<Td width=80 nowrap align=center  >文件編號<br>so van kien</td>		
	</tr>
	<%for x = 1 to pagerec
		if x mod 2 = 0 then 
			wkcolor=""
		else
			wkcolor="PapayaWhip"
		end if 	
	%>
		<Tr height=22 bgcolor="#ffffff" onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"  onclick='chkdata(<%=x-1%>)'>
			<Td align=center><%=(currentpage-1)*pagerec+x%></td>
			<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 16)%></font></td>
			<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 4)%><%=tmpRec(CurrentPage, x, 1)%></font></td>
			<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 2)%></font></td>
			<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 3)%></font></td>
			<Td><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 7)%><BR><%=tmpRec(CurrentPage, x, 8)%></font></td>
			<!--Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 9)%></font></td-->
			<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 13)%></font></td>			
			<Td  width=70  align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 17)%></font></td>
			<Td width=120><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 20)%></font></td>
			<Td width=120 ><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 6)%></font>
				<input type=hidden name=empid value="<%=tmpRec(CurrentPage, x, 3)%>">
				<input type=hidden name=rpno value="<%=tmpRec(CurrentPage, x, 1)%>">
				<input type=hidden name=rptype value="<%=tmpRec(CurrentPage, x, 4)%>">
			</td>
			<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 18)%></font></td>
		</tr>
	<%next%>
</table>	
<input type=hidden name=empid>
<input type=hidden name=rpno>
<input type=hidden name=rptype>
<table width=600 class=txt8>
	<Tr>
	<td align="CENTER" height=40 width=60%>    
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
	<% End If %>
	<input type=button  name=btm class=button value="轉 EXCEL" onclick=goexcel()  >
	</td>		
	<Td height=25 align=center>page : <%=currentpage%>/<%=totalpage%>, recordCount:<%=recordIndB%></td>
	</tr>
</table>	

</form>
</body>
</html>
<script language=vbs> 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function 

function dchg(a) 
	select case a 
		case 1 
			<%=self%>.sortby.value="rpno desc"
		case 2 
			<%=self%>.sortby.value="rp_dat, rpno "
		case 3 
			<%=self%>.sortby.value="a.empid, rp_dat"
		case 4 
			<%=self%>.sortby.value="b.groupid, a.empid, a.rp_dat"
		case 5 
			<%=self%>.sortby.value="rp_func, rpno "
	end select 	
	<%=self%>.totalpage.value="0"
 	<%=self%>.action="<%=self%>.Fore.asp"
 	<%=self%>.submit() 														
end function 
function f()
	<%=self%>.yymm.focus()	
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

function chkdata(index)
	 rpno= <%=self%>.rpno(index).value
	 empid = <%=self%>.empid(index).value
	 rptype= <%=self%>.rptype(index).value
	 open "<%=self%>.foregnd.asp?rpno="&rpno &"&empid="&empid&"&rptype="&rptype, "_self" 
	 
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
function goexcel()
	'open "<%=self%>.toexcel.asp" , "Back" 
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	parent.best.cols="100%,0%"
end function 
</script> 