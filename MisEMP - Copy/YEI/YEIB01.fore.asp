<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%

SELF = "YEIB01"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rds = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

F_whsno = request("F_whsno")
F_groupid = request("F_groupid")
F_zuno = request("F_zuno")
if F_whsno="" then F_whsno="XX"
F_shift=request("F_shift")
F_empid =request("F_empid")
F_country=request("F_country")

sortvalue = request("sortvalue")
if sortvalue ="" then sortvalue="a.country , b.lw, b.lg, a.empid"


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION
khym = request("khym")
if request("khym")="" then
	khym=nowmonth
end if

act = request("act")
khweek = request("khweek")
if khweek="" then khweek="1"

tmw = request("tmw")
if tmw="" then tmw=request("tt")
 '一個月有幾天
cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS

'if khweek="" then khweek=(days\7)

gTotalPage = 1
PageRec = 0    'number of records per page
TableRec = 25    'number of fields per record

Redim tmpRec(gTotalPage, PageRec, TableRec)    

khyears=request("khyears")
F_whsno = request("F_whsno")
f_lev = request("f_lev")
F_country = request("F_country")

if request("khyears")<>"" then  
	sql="select * from KHBFormM_set where years='"& khyears &"' and ( country like '%"&f_country&"%' and whsno like '%"&F_whsno &"%' "&_
			"and isnull(job,'') like '%"&f_lev&"%' ) "&_
			"order by khbid "
	rds.open sql, conn, 3, 3	
	if rds.eof then 
		showstr="N"
	else
		showstr="N"
	end if 		
else	
	sql="select * from KHBFormM_set where years='x' order by khbid "
	set rds=conn.execute(Sql) 	
	showstr="N"
end if 	

act=request("act")
main_khid = request("main_khid")
if main_khid<>"" then 
	sql="select * from KHBFormd_set where khbid='"& main_khid &"'  order by sttno "
	CurrentPage = 1
	rs.Open SQL, conn, 1, 3
	IF NOT RS.EOF THEN
		PageRec = rs.RecordCount
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then
				tmpRec(i, j, 0) = rs("khbid")
				tmpRec(i, j, 1) = rs("sttno")
				tmpRec(i, j, 2) = rs("grade")
				tmpRec(i, j, 3) = rs("fensu")
				tmpRec(i, j, 4) = rs("khstr_cn")
				tmpRec(i, j, 5) = rs("khstr_vn")
				tmpRec(i, j, 6) = rs("aid")
				
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
end if				
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f() 
	<%=self%>.khyears.focus()
	<%=self%>.khyears.select() 
end function 

</SCRIPT>
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
<input name=act value="" type=hidden >
<input name=main_khid value="" type=hidden >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table class="table-borderless table-sm text-secondary txt">
					<TR height=22 >
						<TD align=right nowrap>適用年度</TD>
						<td colspan=6>
							<table border=0 class="table-borderless table-sm text-secondary txt" cellpadding=0 cellspacing=0>
								<tr>
									<td><input type="text" style="width:80px" name=khyears value="<%=left(khym,4)%>" size=6 maxlength=4   ></td>
									<td><input type="text" name="ud"  size=2  style="width:120px"></td>
									<td>
										[請填寫U(上)或D(下) , 未填寫表示整年度適用] &nbsp;&nbsp;&nbsp;&nbsp; [ 新增考核表 ]
									</td>
								</tr>
							</table>										 
						</td>		
					</tr>
					<tr>
						<td align=right nowrap>國籍<BR>Quoc tich</td>
						<td>
							<select name=F_country style="width:180px">
								<%
								if session("rights")<>"" then
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_TYPE  desc"%>
									<option value="" selected >全部(Toan bo) </option>
								<%
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  and sys_type not in ('Tw' ) ORDER BY SYS_TYPE desc "
								end if
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_country then%>selected<%end if%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								SET RST=NOTHING
								%>
							</SELECT>
						</td>
						<TD align=right nowrap>廠別<br>Xuong</TD>
						<td>
							<select name=F_whsno style="width:180px">
								<%
								if session("rights")="0" then
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
									<option value="" selected >全部(Toan bo) </option>
								<%
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("mywhsno") &"' ORDER BY SYS_TYPE "
								end if
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								SET RST=NOTHING
								%>
							</SELECT>
						</td>									
						<TD align=right nowrap>職務<br>Bo Phan</TD>
						<td>
							<select name=f_lev  style="width:180px">
								<%if Session("RIGHTS")<="1"  then%>
									<option value="">全部(Toan bo) </option>
									<%
										SQL="SELECT * FROM BASICCODE WHERE FUNC='lev' and sys_type <>'AAA' ORDER BY SYS_TYPE "
									else
										SQL="SELECT * FROM BASICCODE WHERE FUNC='lev' and sys_type <>'AAA' and  sys_type= '"& session("NETG1") &"' ORDER BY SYS_TYPE "
									end if
									SET RST = CONN.EXECUTE(SQL)
									RESPONSE.WRITE SQL
									WHILE NOT RST.EOF
								%>
									<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
								<%
									RST.MOVENEXT
									WEND
								%>
								<%SET RST=NOTHING %>
							</SELECT>
						</td> 
						<td><input name="btn" value="(S)K.Tra" class="btn btn-sm btn-outline-secondary" type="button" onclick="gosch()"></td>
					</TR>
				</table> 
			</td>
		</tr>
		<tr>
			<td>
				<%if not rds.eof then %>
				<table class="table-borderless table-sm text-secondary">
					<tr>
						<%while not rds.eof  
						  zz = zz+1 
						%>
							<td  align="center" valign="top">
								<a href="vbscript:khbDet(<%=zz-1%>)"><font color="blue"><%=rds("khbid")%></font></a><br><%=rds("khb_memos")%>
								<input name="khbid_m" value="<%=rds("khbid")%>" type="hidden" >
							</td>
						<%rds.movenext
						wend
						%> 
					</tr>
				</table>
				<%end if %>
				<%set rds=nothing %> 
				<input name="khbid_m" value="" type="hidden" >
			</td>
		</tr>
		<tr>
			<td>
				<%if act="Y" then %>
				<table class="table-borderless table-sm text-secondary">
					<tr bgcolor="#FFFF87">
							<td width=40 nowrap align="center">項次</td>
							<td colspan=5>評核主項目</td>				
					</tr>
					<tr>
							<td width=40 nowrap></td>
							<td width=30 nowrap  bgcolor="#CFF3CB">項次</td>
							<td width=70 nowrap bgcolor="#CFF3CB">評語</td>
							<td width=45 nowrap bgcolor="#CFF3CB">分數</td>
							<td  bgcolor="#CFF3CB">評核項目(中)</td>		
					</tr>
				<%for x = 1 to PageRec 
					m_title = tmprec(1,x,1) 
					if len(m_title)=3 then 	
						M_stt = mid("一二三四五六七八九十", right(m_title,1) , 1)  
					else
						m_stt=""
					end if 	

				grade=tmprec(1,x,2) 
				fensu=tmprec(1,x,3) 
				khstr_cn=tmprec(1,x,4) 
				khstr_vn=tmprec(1,x,5)  	
						
				%>
				<%if len(m_title)=3 then %>
					<tr bgcolor="#FFFF87">
						<Td align="center" valign="top"><%=m_stt%></td>
						<td colspan=5><%=khstr_cn%><br><%=khstr_vn%>
							<input type="hidden"  name="M_khstr_CN" class="inputbox8" size=50 >
							<input type="hidden" name="M_khstr_VN" class="inputbox8" size=50 >
						</td>	  
					</tr>
				<%end if %>
				<%if len(m_title)=5 then 
				grade=tmprec(1,x,2) 
				fensu=tmprec(1,x,3) 
				khstr_cn=tmprec(1,x,4) 
				khstr_vn=tmprec(1,x,5) 
				%>
					<tr>
						<Td></td>
						<Td align="center" bgcolor="#CFF3CB" >
							<%=right(m_title,2)%>
						</td>
						<td bgcolor="#CFF3CB"  ><%=grade%></td>	
						<td bgcolor="#CFF3CB" align="center"><%=fensu%></td>	
						<td " bgcolor="#CFF3CB" ><%=khstr_cn%><br><%=khstr_vn%></td>	
					</tr>
				<%end if %>
					<input name="sttno" value="<%=tmprec(1,x,1)%>" type="hidden" >
				<%next%>
				</table>
				<%end if%>
			</td>
		</tr>
	</table>
			
</form>


</body>
</html>
<script language=vbscript>
function khbDet(index)
	if <%=self%>.khbid_m(index).value<>"" then 
		<%=self%>.act.value="Y"
		<%=self%>.main_khid.value=<%=self%>.khbid_m(index).value
		<%=self%>.action="<%=self%>.fore.asp"
		<%=self%>.submit()
	end if 	
end function  

function gosch()
	<%=self%>.action="<%=self%>.fore.asp"
	<%=self%>.submit()
end function
</script>

