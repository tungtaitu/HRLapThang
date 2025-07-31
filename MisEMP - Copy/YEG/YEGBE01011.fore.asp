<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
SELF = "YEGBE0101"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")  
Set rst = Server.CreateObject("ADODB.Recordset")  

gTotalPage = 1
PageRec = 32    'number of records per page
TableRec = 10    'number of fields per record  

'Response.Write nowmonth &"<BR>"
'Response.Write calcmonth &"<BR>"      
'Response.End 
'a=4.35689 
'b = - Int(-a)
'response.write b &<BR>" 

groupid = request("groupid")
shift = request("shift")
zuno = request("zuno")
calcH = request("calcH")
allstr =request("allstr")

yymm = request("yymm")  
if yymm<>"" then 
	chkym=left(yymm,4)&"/"&right(yymm,2)
	cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01") 
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
	
	sql="select * from YDBMCALE where convert(char(6),dat,112)='"& yymm &"' order by dat "
	rst.open sql, conn, 3, 3 
	if rst.eof then 
		for k = 1 to days 
			insdate = chkym & "/"&right("00"&k,2)
			if weekday(insdate)=1 then 
				stas="H2" 
			else
				stas="H1"
			end if 		
			sql =" insert into YDBMCALE (dat, status ) values ('"& insdate &"' , '"& stas &"' ) "
			'response.write sql &"<BR>"
			conn.execute(sql)
			'response.write sql 
		next 			
	end if 
	set rst=nothing 
	'----------------------------------------------------------------------------------------------
	if request("TotalPage") = "" or request("TotalPage") = "0" then 
		CurrentPage = 1	
		sqlstr = "select * from YDBMCALE where convert(char(6),dat,112)='"& yymm &"' order by dat "
		rs.Open SQLstr, conn, 3, 3 
		IF NOT RS.EOF THEN 
			rs.PageSize = PageRec 
			RecordInDB = rs.RecordCount 
			TotalPage = rs.PageCount  
			gTotalPage = TotalPage
		END IF 	 
	
		Redim tmpRec(TotalPage, PageRec, TableRec)   'Array
		
		for i = 1 to TotalPage 
			for j = 1 to PageRec
				if not rs.EOF then 
					for k=1 to TableRec-1
						tmpRec(i, j, 0) = "no"
						tmpRec(i, j, 1) = year(trim(rs("dat")))&"/"&right("00"&month(rs("dat")),2)&"/"&right("00"&day(rs("dat")),2)
						tmpRec(i, j, 2) = trim(rs("status"))										
						tmpRec(i, j, 3)= mid("日一二三四五六",weekday(tmpRec(i, j, 1)) , 1 )
						
					next
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
		'Session("EMPBASICB") = tmpRec	 
	end if 	 
else
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 
end if 

  
 
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	<%=self%>.Uptim(0).focus()	
	<%=self%>.Uptim(0).select()
end function  

function checkall()
	if <%=self%>.all.checked then 
		<%=self%>.allstr.value="Y"
		if <%=self%>.days.value>"0" and  <%=self%>.days.value<>"" then 
			for z=1 to  <%=self%>.days.value 
				<%=self%>.Uptim(z-1).value="09:00"
			next
		end if 
	else
		<%=self%>.allstr.value=""
		if <%=self%>.days.value>"0" and  <%=self%>.days.value<>"" then 
			for z=1 to  <%=self%>.days.value 
				<%=self%>.Uptim(z-1).value=""
			next
		end if 	
	end if	
end function 
-->
</SCRIPT>  
</head>   
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<INPUT TYPE=HIDDEN NAME="workdays" VALUE="<%=days%>">

<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
        <img border="0" src="../image/icon.gif" align="absmiddle"> 
        <%=session("pgname")%> </TD>	 
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<table width=460><tr><td>
	<table width=480 align=center class=font9>
		<tr height=30>
			<td width=70 align=right>設定年月：</td>
			<td  colspan=3>
				<select name=yymm class=font9 onchange="datachg()">
					<option value=""> </option>
					<%for z = 1 to 12 
					  yymmvalue = year(date())&right("00"&z,2)
					%>
						<option value="<%=yymmvalue%>" <%if yymmvalue=yymm then %>selected<%end if%>><%=yymmvalue%></option>
					<%next%>	
				</select>  				
				<input class=readonly readonly  name=days value="<%=days%>" size=5>　　　				
			</td> 			
			<td width=60 align=right></td> 
			<td>
				
			</td> 			
		</tr>		
		<tr>
			<td width=70 align=right nowrap>單位：</td>
			<td width=80>
				<select name=GROUPID  class=font9  onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'   and sys_type <>'AAA'and sys_type>='A033' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=GROUPID THEN %> SELECTED <%END IF%> >
				<%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			</td>
			<td width=60 align=right>組別：</td>
			<td width=80>
				<select name=Zuno  class=font9  >
				<%IF GROUPID="" THEN %><option value="">-----</option><%END IF%>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='ZUNO' and left(sys_type,4) like '%"& groupid &"' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=Zuno THEN %> SELECTED <%END IF%> >
				<%=right(RST("SYS_TYPE"),1)%>-<%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
				</SELECT>
				<%SET RST=NOTHING %>		
			</td>			
			<td width=60 align=right>班別：</td>
			<td>
				<select name=shift  class=font9  >
				<option value="">---</option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='shift' ORDER BY SYS_value "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=shift THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			</td>
			
		</tr>
	</table>	

	<table width=450 class=font9 border=0 align=center>
		<tr bgcolor=#cccccc height=30>
			<td width=90 align=center height=35 >日期</td>
			<td width=40 align=center>星期</td>			
			<td width=80 align=center>上班時間</td>			
			<td width=5 align=center bgcolor=#ffffff>
			<td width=90 align=center>日期</td>
			<td width=40 align=center>星期</td>			
			<td width=80 align=center>上班時間</td>			
			<td width=5 align=center bgcolor=#ffffff>
		</tr>
		<%for x = 1 to days 				
		%>
		<%if x mod 2 = 1 then %><tr bgcolor="Beige"><%end if %>
			<td align=center>
				<%'response.write x  %>
				<%if weekday(tmprec(1,x,1))=1 then %>
					<font color=red><%=tmprec(1,x,1)%></font>
				<%else%>	
					<font color=black><%=tmprec(1,x,1)%></font>
				<%end if%>
				<%'=tmprec(1,x,4)%>
				<input type=hidden name=dat  size=2 value=<%=tmprec(1,x,1)%> class=inputbox>
				<input type=hidden name=b_sts  size=2 value=<%=tmprec(1,x,2)%> class=inputbox>
				<input type=hidden name=status  size=2 value=<%=tmprec(1,x,2)%>  class=inputbox >
			</td>
			<td align=center><%=tmprec(1,x,3)%></td>					 
			<td align=center>
				<input name=Uptim  size=10 class=inputbox value="" style="text-align:center"  onblur="timchg(<%=x-1%>)" maxlength=5>
			</td>			
			<td width=5 bgcolor=#ffffff >	
		<%if x mod 2 = 0 then %></tr><%end if %>
		<%next%>
	</table>	
	<br>
	<TABLE WIDTH=400>
	<tr  >
		<TD align=center>				 
		<%if UCASE(session("mode"))="W" then%>
		<input type="button" name=send value="確　　認"  class=button onclick="go()" onkeydown="go()" > 
		<input type=RESET name=send value="取　　消"  class=button>		
		<%end if%>
		</TD>
	</TR>
</TABLE>
</form>
</td></tr></table> 

</body>
</html> 

<script language=vbs> 
function timchg(index)
	days=<%=self%>.workdays.value
	if days<>"" and days>"0"  then 
		if index>0 then 
			if <%=self%>.Uptim(index-1).value<>"" and <%=self%>.status(index).value="H1" then
				<%=self%>.Uptim(index).value=<%=self%>.Uptim(index-1).value
			elseif <%=self%>.status(index).value<>"H1" then 
				<%=self%>.Uptim(index).value=""
			else	
				if <%=self%>.uptim(index).value<>"" then 
					<%=self%>.Uptim(index).value=left(<%=self%>.Uptim(index).value,2)&":"&right(<%=self%>.Uptim(index).value,2)	
				end if 	
			end if				
		else
			if <%=self%>.uptim(index).value<>"" then 
				<%=self%>.Uptim(index).value=left(<%=self%>.Uptim(index).value,2)&":"&right(<%=self%>.Uptim(index).value,2)
			end if
		end if 
	end if 
	
end function 

function BACKMAIN() 	
	open "../main.asp" , "_self"
end function 

function datachg()
	<%=self%>.action = "<%=self%>.fore.asp"
	<%=self%>.submit()
end function  

function t2chg(index)
	if <%=self%>.t2(index).checked=true then 
		if <%=self%>.t3(index).checked=true then 
			<%=self%>.t3(index).checked=false 
		end if 
		<%=self%>.status(index).value="H2"
	else
		<%=self%>.status(index).value=<%=self%>.b_sts(index).value 
		if <%=self%>.b_sts(index).value="H3" then 
			<%=self%>.t3(index).checked=true 
		end if 	
	end if 	 
end function 

function t3chg(index)
	if <%=self%>.t3(index).checked=true then 
		if <%=self%>.t2(index).checked=true then 
			<%=self%>.t2(index).checked=false 
		end if 
		<%=self%>.status(index).value="H3"
	else
		<%=self%>.status(index).value=<%=self%>.b_sts(index).value 
		if <%=self%>.b_sts(index).value="H2" then 
			<%=self%>.t2(index).checked=true 
		end if 	
	end if 	 
end function 

function go()
	if <%=self%>.yymm.value="" then 
		alert "請選擇年月!!"
		<%=self%>.yymm.focus()
		exit function
	end if 
	if <%=self%>.groupid.value="" then 
		alert "請選擇單位!!"
		<%=self%>.groupid.focus()
		exit function
	end if 	
	if <%=self%>.shift.value="" then 
		alert "請選擇班別!!"
		<%=self%>.shift.focus()
		exit function
	end if 
	<%=self%>.action = "<%=self%>.upd.asp"
	<%=self%>.submit()
end function 
 
 
	
</script>

