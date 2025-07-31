<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

SELF = "YEIE0201"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rds = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 
 
'一個月有幾天 
'cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01") 
'days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
'ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS   


khyears=request("khyears")
khud=Ucase(Trim(request("khud")))
F_whsno = request("F_whsno")
F_groupid = request("F_groupid")
F_zuno = request("F_zuno") 
F_shift=request("F_shift")
F_empid =request("f_empid")
F_country=request("F_country") 

f_khBid=khyears&khud  
 
if  khud="U" then 
	d1_str = khyears&"0101"
	d2_str= khyears&"0630" 
elseif khud="D" then  
	d1_str = khyears&"0701"
	d2_str= khyears&"1231" 
end if 	 

'if khweek="" then khweek=(days\7)    


gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 40   'number of fields per record    
 
sql="select * from  fn_yeie0201  ( '"&d1_str&"','"&d2_str&"','"&khyears&"','"&khud&"','"&F_whsno&"','"&F_groupid&"',  "&_
		"'"&F_country&"','"&F_empid &"' ) "&_
		"order by groupid, empid " 	
 
'response.write sql 	
response.end 			
if request("TotalPage") = "" or request("TotalPage") = "0" then 

	CurrentPage = 1	
	
	rds.Open SQL, conn, 1, 3 
	
	IF NOT RdS.EOF THEN 	
		PageRec = rds.RecordCount 
		rds.PageSize = PageRec 
		RecordInDB = rds.RecordCount 
		TotalPage = rds.PageCount  
		gTotalPage = TotalPage
	END IF 	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rds.EOF then 			
				tmpRec(i, j, 0) = rds("EMPkhid")
				tmpRec(i, j, 1) = rds("EMPID")
				tmpRec(i, j, 2) = rds("empnam_cn")
				tmpRec(i, j, 3) = rds("empnam_vn")
				tmpRec(i, j, 4) = rds("nindat") 
				tmpRec(i, j, 5) = rds("outdate") 
				tmpRec(i, j, 6) = rds("gstr")
				tmpRec(i, j, 7) = rds("zstr") 
				tmpRec(i, j, 8) = rds("country")  
				tmpRec(i, j, 9) = rds("whsno") 
				tmpRec(i, j, 10) = rds("groupid")
				tmpRec(i, j, 11) = rds("zuno") 
				tmpRec(i, j, 12) = rds("shift")
				tmpRec(i, j, 13) = rds("fensu")
				tmpRec(i, j, 14) = rds("grade")
				tmpRec(i, j, 15) = rds("zcqid")
				tmpRec(i, j, 16) = rds("zcqname")
				tmpRec(i, j, 17) = rds("jcqid")
				tmpRec(i, j, 18) = rds("jcqname")
				tmpRec(i, j, 19) = rds("hcqid")
				tmpRec(i, j, 20) = rds("hcqname")
				tmpRec(i, j, 21) = rds("khporc_sts")
				tmpRec(i, j, 22) = rds("formKhbid")
				tmpRec(i, j, 23) = rds("job")
				tmpRec(i, j, 24) = rds("jstr")
				tmpRec(i, j, 25) = rds("kzhour")
				tmpRec(i, j, 26) = rds("flz")
				tmpRec(i, j, 27) = rds("jiaA")
				tmpRec(i, j, 28) = rds("jiaB")
				tmpRec(i, j, 29) = rds("z_fs")
				tmpRec(i, j, 30) = rds("z_kj")
				tmpRec(i, j, 31) = rds("j_fs")
				tmpRec(i, j, 32) = rds("j_kj")
				tmpRec(i, j, 33) = rds("h_fs")
				tmpRec(i, j, 34) = rds("h_kj") 				
				rds.MoveNext 
			else 
				exit for 
			end if 
	 	next 
	
	 	if rds.EOF then 
			rds.Close 
			Set rds = nothing
			exit for 
	 	end if   
	next	
end if
session("yeie0201B") = tmprec

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 	
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
 
end function
 

function datachg() 
	<%=self%>.totalpage.value="0"
	<%=self%>.sortvalue.value=""
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
end function   

function datachg2() 
	<%=self%>.totalpage.value="0"	
	<%=self%>.act.value="A"
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
end function 

function sortby(a)
	if a=1 then 
		<%=self%>.sortvalue.value="a.khz, a.empid"
	elseif a=2 then	
		<%=self%>.sortvalue.value="a.empid"
	elseif a=3 then	
		<%=self%>.sortvalue.value="b.nindat, a.empid"
	elseif a=4 then	
		<%=self%>.sortvalue.value="a.monthfen desc, a.empid"
	elseif a=5 then	
		<%=self%>.sortvalue.value="len(a.khs) desc, a.khs, a.khz, a.empid"
	else
		<%=self%>.sortvalue.value="b.country , a.khw, a.khg, a.empid"			
	end if 	 
	<%=self%>.totalpage.value="0"
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
	'alert a 
end function 

 
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form  name="<%=self%>" method="post" action="<%=self%>.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=hidden NAME=days VALUE="<%=days%>">
<INPUT TYPE=hidden NAME=sortvalue VALUE="<%=sortvalue%>">
<input name=act value="<%=act%>" type=hidden >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		<%=session("pgname")%>
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<TABLE WIDTH=650 CLASS=TXT BORDER=0>
	<TR height=22 >		
		<TD align=right>考核年度</TD>
		<td colspan=6>
			<input class=inputbox name=khyearsUd value="<%=khyears&khud%>" size=6 maxlength=6  >
			<input type="hidden" name=khyears value="<%=khyears%>" size=6 maxlength=6  >
			<input type="hidden" name=khud value="<%=khud%>" size=6 maxlength=6  > 
		</td> 
	</tr>
	<tr  height=22>	
		<TD align=right>廠別<br>Xuong</TD>
		<td>
			<select name=F_whsno class=txt8 style='width:100' >					
					<%
					if session("rights")="0" then 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
						<option value="" selected >全部(Toan bo) </option>
					<%		
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE " 'and sys_type='"& session("NETWHSNO") &"'
					end if	
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>	
		</td>
		<td align=right>國籍<BR>Quoc</td>
		<td>
			<select name=F_country class=txt8 style='width:70' onchange="datachg()">									
					<option value="" selected >全部(Toan bo) </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  and sys_type not in ('Tw' ) ORDER BY SYS_TYPE desc "					
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_country then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>	
		</td>		
		
		<TD align=right>部門<br>Bo Phan</TD>
		<td>
			<select name=F_groupid  class=txt8  style='width:80' onchange="datachg()">			
					<option value="">全部(Toan bo) </option>
					<% 
					SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE " 
					SET RST = CONN.EXECUTE(SQL)
					RESPONSE.WRITE SQL 
					WHILE NOT RST.EOF  
				%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
				<%
					RST.MOVENEXT
					WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
			<select name=F_zuno  class=txt8 style='width:70'   onchange="datachg()" >				
				<option value=""></option>
				<%
					SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and sys_type <>'XX' and  left(sys_type,4)= '"& F_groupid &"' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					RESPONSE.WRITE SQL 
					WHILE NOT RST.EOF  
				%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_zuno then%>selected<%end if%>><%=right(RST("SYS_TYPE"),1)%>-<%=RST("SYS_VALUE")%></option>				 
				<%
					RST.MOVENEXT
					WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>	
		</td>
		<td align=right>班別<BR>Ca</td>
		<td>
			<select name=F_shift  class=txt8 onchange="datachg()"  >
				<option value=""></option>
				<option value="ALL" <%if F_shift="ALL" then%>selected<%end if%>>日</option>
				<option value="A" <%if F_shift="A" then%>selected<%end if%>>A班</option>
				<option value="B" <%if F_shift="B" then%>selected<%end if%>>B班</option>
			</select>
			<input type=button name=send value="(S)查詢" class=button onclick="datachg()">
		</td>  	 
	</TR>	 
</table>
 
<hr size=0	style='border: 1px dotted #999999;' align=left width=600> 

<table BORDER=0 cellspacing="1" cellpadding="1"  class=txt8 >
	<tr height=22 bgcolor=#e4e4e4>
		<Td width=20 nowrap align=center rowspan=3>STT</td>
		<Td width=45 nowrap align=center rowspan=3>部門</td>		
		<Td width=50 nowrap align=center rowspan=3>單位<br></td>		
		<Td width=50 nowrap align=center rowspan=3>班別<br></td>		
		<Td width=30 nowrap  align=center rowspan=3>評核<br></td>
		<Td width=50 nowrap align=center rowspan=3>工號<br></td>
		<Td width=100 nowrap align=center  rowspan=3>姓名</td>
		<Td width=65 nowrap  align=center rowspan=3>到職日<br></td>		
		<Td width=40 nowrap  align=center rowspan=3 style="display:none">分數<br></td>
		<Td width=40 nowrap  align=center rowspan=3 style="display:none">考績<br></td>
		<Td   nowrap align=center  colspan=6>簽核程序</td> 
	</tr>  
	<tr bgcolor=#e4e4e4>
		<td width=60 nowrap colspan=2>直接主管</td>
		<td width=60 nowrap colspan=2>間接主管</td>
		<td width=60 nowrap colspan=2>核決主管</td>
	</tr>
	<tr  bgcolor=#e4e4e4>
		<td width=30 nowrap>分數</td>
		<td width=30 nowrap>考績</td>
		<td width=30 nowrap>分數</td>
		<td width=30 nowrap>考績</td>
		<td width=30 nowrap>分數</td>
		<td width=30 nowrap>考績</td>
	</tr>
	<%
	for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR="#DFEFFF"
		END IF 	 		
	%>
	 
	<TR BGCOLOR="<%=WKCOLOR%>" > 	 
		<Td align=center ><%=CurrentRow%></td>
		<Td align=center ><%=tmpRec(CurrentPage, CurrentRow, 6)%></td>
		<td align=left ><%=tmpRec(CurrentPage, CurrentRow, 7)%></td>
		<Td align=center          ><%=tmpRec(CurrentPage, CurrentRow, 12)%></td> <!--shift-->
		<td>
			<input name=khyn class="inputbox8" size=2 style="text-align:center;border: 1px solid <%=WKCOLOR%> ; background-color:<%=WKCOLOR%>" readonly ></td>
		<Td align=center   >
			<%=tmpRec(CurrentPage, CurrentRow, 1)%>
			<input type=hidden name=empid  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
			<input type=hidden name=khbid  value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>">
			<input type=hidden name=empkhid  value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>">
		</td>
		<Td ><a href="vbscript:khemp(<%=CurrentRow-1%>)">
			<%=tmpRec(CurrentPage, CurrentRow, 2)%><br><%=left(tmpRec(CurrentPage, CurrentRow, 3),15)%>
			</a>
		</td>		
		<Td align=center title='紅色表示員工離職日'><%=tmpRec(CurrentPage, CurrentRow, 4)%><br><font color=red>
			<%=tmpRec(CurrentPage, CurrentRow, 5)%></font>
		</td> 
		<td style="display:none"><input name=funsu class="inputbox8" size=3 style="text-align:center;"></td>
		<td style="display:none"><input name=grade class="inputbox8" size=3 style="text-align:center;"></td>
		<td align=center colspan=2>
			<%=tmpRec(CurrentPage, CurrentRow, 16)%>
			<br>
			<input name="zfs" class="inputbox8" size=3>
			<input name="zfs" class="inputbox8" size=3>
		</td>  		
		<td align=center colspan=2>
			<%=tmpRec(CurrentPage, CurrentRow, 18)%>
			<br>
			<input name="zfs" class="inputbox8" size=3>
			<input name="zfs" class="inputbox8" size=3>
		</td>  
		<td align=center colspan=2>
			<%=tmpRec(CurrentPage, CurrentRow, 20)%><br>
			<input name="zfs" class="inputbox8" size=3>
			<input name="zfs" class="inputbox8" size=3>
		</td>  
	</tr>				
	<%next%> 
	<input type=hidden name=empid  value="">	
	<input type=hidden name=khbid  value="">	
	<input type=hidden name=empkhid  value="">	
</table>

<TABLE WIDTH=600>
	<tr ALIGN=center>
	<TD >
		<input type=button  name=send value="(M)回主畫面" class=button onclick=backM()>
		<input type=button  name=send value="下載到Excel" class=button onclick=goexcel()>
	</TD>
	</TR>
</TABLE>


</form>


</body>
</html>
<script language=vbscript>  

function khemp(index)
	empidstr = <%=self%>.empid(index).value	
	khyearud = <%=self%>.khyearsUd.value 
	empkhid= <%=self%>.empkhid(index).value	
	khbid= <%=self%>.khbid(index).value	
	open "<%=self%>B.Foregnd.asp?index="&index &"&empid="& empidstr &"&khyear="& khyearud &"&empkhid="& empkhid &"&khbid="& khbid , "Back" 
	parent.best.cols="0%,100%"
end function  

 
function  backM()	
	open "<%=self%>.asp", "_self"
	
end function 
  
</script>

