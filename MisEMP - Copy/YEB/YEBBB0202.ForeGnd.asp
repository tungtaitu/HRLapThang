<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->

<%
'on error resume next 
session.codepage="65001"
SELF = "YEBBB0202"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
job = trim(request("job"))
country = request("country")
QUERYX = trim(request("empid1"))
outemp = request("outemp")

shift = request("shift")

yymm = request("yymm")

F_name = request("F_name")
F_EMPID = request("F_EMPID")
F_indat = request("F_indat")
F_outdat = request("F_outdat")
BirDay = request("BirDay")
AGES = request("AGES")
F_cstr = request("F_cstr")
F_wstr = request("F_wstr")
F_ustr = request("F_ustr")
F_gstr = request("F_gstr")
F_zstr = request("F_zstr")
F_shift = request("F_shift")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 35    'number of fields per record
NOWMONTH=CSTR(YEAR(DATE()))&RIGHT("00"&CSTR(MONTH(DATE())),2)
'if RIGHT("00"&CSTR(MONTH(DATE())),2)="12" then 
'	NOWMONTH=CSTR(YEAR(DATE())+1)&"01"
'else	
'	NOWMONTH=CSTR(YEAR(DATE()))&RIGHT("00"&CSTR(MONTH(DATE())+1),2)
'end if 

NOWDAY=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)


sql="select c.cym,  a.yymm, a.Cwhsno , a.Cgroupid, a.Czuno, a.Cshift , a.Cmemo, b.* from  "&_
	"( select *from view_empfile  where empid ='" & QUERYX &"'  ) b  "&_
	"left join (  "&_
 	"Select convert(varchar(6),dat,112) cym  from ydbmcale group by convert(varchar(6),dat,112) ) c on c.cym >=b.inyymm  "&_
	"left join (  "&_
	"select yymm, empid, whsno Cwhsno , groupid Cgroupid , zuno Czuno , shift Cshift , memo as Cmemo from bempg  "&_
	"group by yymm , empid, whsno , groupid, zuno, shift , memo  "&_
	") a on b.empid = a.empid  and a.yymm =c.cym "&_
	"where c.cym <=convert(char(6), getdate()+30,112) "&_
	"order by c.cym desc " 

	
'response.write sql 
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 1, 3
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
			F_name=trim(rs("empnam_cn"))&trim(rs("empnam_vn"))
			F_empid = rs("empid")
			F_empnam_cn = trim(rs("empnam_cn"))
			F_empnam_vn = trim(rs("empnam_vn")) 						
			F_cstr = RS("cstr")			
			birDay = rs("BYY")&"/"&right("00"&rs("BMM"),2)&"/"&right("00"&rs("Bdd"),2)
			ages  = rs("ages")
			sex = rs("sex")
			F_indat = rs("nindat")
			F_outdat = rs("outdate") 

			F_wstr = rs("wstr")
			F_ustr = RS("ustr")
			F_cstr = RS("cstr")
			F_gstr = RS("gstr")
			F_zstr = RS("zstr") 
			F_shift = RS("sstr") 
			F_BHDAt = RS("bhdat")
			country=rs("country")
			autoid = rs("autoid")
			
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("empid"))
			tmpRec(i, j, 2) = trim(rs("empnam_cn"))
			tmpRec(i, j, 3) = trim(rs("empnam_vn"))
			tmpRec(i, j, 4) = rs("country")
			tmpRec(i, j, 5) = rs("nindat")
			tmpRec(i, j, 6) = rs("Job")
			tmpRec(i, j, 7) = rs("cwhsno")
			tmpRec(i, j, 8) = rs("unitno")
			tmpRec(i, j, 9)	=RS("cgroupid")
			tmpRec(i, j, 10)=RS("czuno")
			tmpRec(i, j, 11)=RS("wstr")
			tmpRec(i, j, 12)=RS("ustr")
			tmpRec(i, j, 13)=RS("gstr")
			tmpRec(i, j, 14)=RS("zstr")
			tmpRec(i, j, 15)=RS("jstr")
			tmpRec(i, j, 16)=RS("cstr")
			tmpRec(i, j, 17)=RS("autoid")
			tmpRec(i, j, 18)= "" 'RS("aid")
			tmpRec(i, j, 19)=RS("bhdat")
			tmpRec(i, j, 20)=RS("outdate")
			tmpRec(i, j, 21)=RS("cSHIFT")
			tmpRec(i, j, 22)=RS("GTDAT")
			tmpRec(i, j, 23)=rs("Cmemo") 			
			tmpRec(i, j, 25)=rs("cym")
			if rs("cym") < nowmonth then 
				tmpRec(i, j, 26)="DarkGray"
			elseif rs("cym") = nowmonth then 
				tmpRec(i, j, 26)="Blue"
			else
				tmpRec(i, j, 26)="Firebrick"
			end if 
			tmpRec(i, j, 27)=left(replace(RS("outdate"),"/",""),6)
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
	Session("YEBBB0202") = tmpRec
else 	
	TotalPage = cint(request("TotalPage"))
	gTotalPage = cint(request("gTotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBBB0202")

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
<body  marginheight="0"   >
<form name="<%=self%>" method="post" action="<%=self%>.foreGnd.asp">
	<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
	<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
	<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
	<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
	<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
	<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
	<INPUT TYPE=hidden NAME=NowMonth VALUE="<%=NowMonth%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%">
		<tr>
			<td>
				<table class="table-borderless table-sm text-secondary txt">
					<TR>
						<TD align="right">工號/姓名<br>So the/Ho Ten</TD>
						<TD colspan=5 valign="top">
							<a href='javascript:oktest(<%=autoid%>)'>
							<%=F_EMPID%>&nbsp;<%=F_name%> (<%=F_cstr%> )</a>
							<INPUT name="f_empid" SIZE=12 CLASS=READONLY VALUE="<%=F_EMPID%>" READONLY  type="hidden" > 
							<INPUT NAME=F_name SIZE=25 CLASS=readonly readonly  VALUE="<%=F_name%>"   type="hidden"  >
							<input type="hidden" name="ct" value="<%=country%>">	
						</TD> 
					</TR>
					<tr>
						<td nowrap align=right>目前單位<br>Xuong/bo phan</td>
						<td colspan=3 valign="top">(<%=F_wstr%>)&nbsp;<%=F_ustr%>-<%=F_gstr%>-<%=F_zstr%> ( ca:<%=F_shift%> )
							<input name=F_wstr size=10 class='readonly' readonly value="<%=F_wstr%>" type="hidden" >  
							<input name=F_ustr size=10 class='readonly' readonly value="<%=F_ustr%>" type="hidden" > 
							<input name=F_gstr size=5 class='readonly' readonly value="<%=F_gstr%>" type="hidden" > 
							<input name=F_zstr size=5 class='readonly' readonly value="<%=F_zstr%>" type="hidden" > 
							<input name=F_shift size=4 class='readonly' readonly value="<%=F_shift%>"  type="hidden" > 
						</td>
						<TD align=right >到職日<br>NVX</TD>
						<TD valign="top"><%=F_indat%>
							<INPUT NAME=F_indat SIZE=12 CLASS=readonly readonly  VALUE="<%=F_indat%>"  type="hidden"  >			
						</TD>
						<TD  nowrap align=right >離職日<br>NTV</TD>
						<TD valign="top"><%=F_outdat%>
							<INPUT NAME=F_outdat SIZE=12 CLASS=readonly readonly  VALUE="<%=F_outdat%>"   type="hidden" >			
						</TD>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid">	
					<tr class="header">									
						<TD nowrap >異動年月<br>yymm</TD> 		 		
						<TD nowrap>廠別<br>Xuong</TD>
						<TD nowrap colspan=2>單位<br>Bo Phan</TD>
						<TD nowrap>班別<br>Ca</TD>
						<TD nowrap>異動說明<br>ghi chu(Ly do)</TD>		
					</TR>
					<%for CurrentRow = 1 to PageRec
						 
						X_empid = TRIM(tmpRec(CurrentPage, CurrentRow, 1))  '員工編號
						X_outdat = trim(tmpRec(CurrentPage, CurrentRow, 20)) '離職日
						X_OUTYM = left( trim(tmpRec(CurrentPage, CurrentRow, 20)),4)&mid( trim(tmpRec(CurrentPage, CurrentRow, 20)),6,2) '離職年月
						X_CalcYM = TRIM(tmpRec(CurrentPage, CurrentRow, 25)) '異動年月
						'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
					%>
					<TR> 	 
						<TD align=center ><font color="<%=tmpRec(CurrentPage, CurrentRow, 26)%>"><b><%=tmpRec(CurrentPage, CurrentRow, 25)%></b></font></TD> 		
						<TD nowrap>
							<INPUT NAME=op TYPE=HIDDEN size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>">
							<%IF  X_empid<>"" THEN %>
								<%'異動年月<本月
								if  X_calcYM < trim(nowmonth) and session("Netuser")<>"PELIN"  then  %>
									<INPUT NAME=whsno TYPE=HIDDEN value="<%=TRIM(tmpRec(CurrentPage, CurrentRow, 7))%>" >
									<SELECT NAME=whsnoT   disabled >
										<OPTION VALUE="">--</option>
										<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='whsno' ORDER BY SYS_TYPE "
										  SET RDS=CONN.EXECUTE(SQL)
										  WHILE NOT RDS.EOF
										%>
										<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 7)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
										<%RDS.MOVENEXT
										WEND
										rds.close
										SET RDS=NOTHING
										%>
									</SELECT>
								<%'異動年月>離職年月 and 離職日期<>""	
								elseif  X_calcYM>X_OUTYM and X_outdat<>"" then %>
									<INPUT NAME=whsno TYPE=HIDDEN>	
								<%else%>
									<SELECT NAME=whsno  onchange="datachg(<%=currentRow-1%>)">
										<OPTION VALUE="">--</option>
										<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='whsno' ORDER BY SYS_TYPE "
										  SET RDS=CONN.EXECUTE(SQL)
										  WHILE NOT RDS.EOF
										%>
										<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 7)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
										<%RDS.MOVENEXT
										WEND
										rds.close
										SET RDS=NOTHING
										%>
									</SELECT>
								<%end if%>	
							<%ELSE%>
								<INPUT NAME=whsno TYPE=HIDDEN>	
							<%END IF%> 			
						</TD>									
						
							<%IF  X_empid<>"" THEN %>
								<%'異動年月<本月
								if  X_calcYM < trim(nowmonth)  and session("Netuser")<>"PELIN"  then  %>
									<td>
										<INPUT NAME=groupid TYPE=HIDDEN value="<%=TRIM(tmpRec(CurrentPage, CurrentRow, 9))%>">
										<INPUT NAME=zuno TYPE=HIDDEN value="<%=TRIM(tmpRec(CurrentPage, CurrentRow, 10))%>">
										<SELECT NAME=groupidT   disabled>
											<OPTION VALUE="">--</option>
											<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='groupid' ORDER BY SYS_TYPE "
											  SET RDS=CONN.EXECUTE(SQL)
											  WHILE NOT RDS.EOF
											%>
											<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 9)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
											<%RDS.MOVENEXT
											WEND
											SET RDS=NOTHING
											%>
										</SELECT>
									</td>
									<td>
										<SELECT NAME=zunoT   disabled>
											<option value="">------</option>
											<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='zuno' ORDER BY SYS_TYPE "
											  SET RDS=CONN.EXECUTE(SQL)
											  WHILE NOT RDS.EOF
											%>
											<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 10)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
											<%RDS.MOVENEXT
											WEND
											SET RDS=NOTHING
											%>
										</SELECT>
									</td>
								<%'異動年月>離職年月 and 離職日期<>""	
								elseif  X_calcYM>X_OUTYM and X_outdat<>"" then %>
									<td>
										<INPUT NAME=groupid TYPE=HIDDEN>	
										<INPUT NAME=zuno TYPE=HIDDEN>	
									</td>
								<%else%>
									<td>
										<SELECT NAME=groupid   onchange="gchg(<%=currentRow-1%>)">
											<OPTION VALUE="">--</option>
											<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='groupid' ORDER BY SYS_TYPE "
											  SET RDS=CONN.EXECUTE(SQL)
											  WHILE NOT RDS.EOF
											%>
											<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 9)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
											<%RDS.MOVENEXT
											WEND
											SET RDS=NOTHING
											%>
										</SELECT>
									</td>
									<td>
										<SELECT NAME=zuno  onchange="datachg(<%=currentRow-1%>)">
											<OPTION VALUE="">--</option>
											<%if TRIM(tmpRec(CurrentPage, CurrentRow, 10))="" then %><option value="">------</option><%end if%>
											<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='zuno' and left(sys_type,4)='"& TRIM(tmpRec(CurrentPage, CurrentRow, 9)) &"' ORDER BY SYS_TYPE "
											  SET RDS=CONN.EXECUTE(SQL)
											  WHILE NOT RDS.EOF
											%>
											<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 10)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
											<%RDS.MOVENEXT
											WEND
											rds.close
											SET RDS=NOTHING
											%>
										</SELECT>
									</td>
								<%end if%>	
							<%ELSE%>
						<td>
							<INPUT NAME=groupid TYPE=HIDDEN>	
							<INPUT NAME=zuno TYPE=HIDDEN>
						</td>
							<%END IF%>
								
						<TD>
							<%IF  X_empid<>"" THEN %>
								<%'異動年月<本月
								if  X_calcYM < trim(nowmonth)  and session("Netuser")<>"PELIN" then  %>
									<INPUT NAME=shift TYPE=HIDDEN value="<%=TRIM(tmpRec(CurrentPage, CurrentRow, 21))%>">
									<SELECT NAME=shiftT   disabled >
										<OPTION VALUE="">--</option>
										<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='shift' ORDER BY SYS_TYPE "
										  SET RDS=CONN.EXECUTE(SQL)
										  WHILE NOT RDS.EOF
										%>
										<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 21)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
										<%RDS.MOVENEXT
										WEND
										rds.close
										SET RDS=NOTHING
										%>
									</SELECT>
								<%'異動年月>離職年月 and 離職日期<>""	
								elseif  X_calcYM>X_OUTYM and X_outdat<>"" then %>
									<INPUT NAME=shift TYPE=HIDDEN>	
								<%else%>
									<SELECT NAME=shift  onchange="datachg(<%=currentRow-1%>)">
										<OPTION VALUE="">--</option>
										<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='shift' ORDER BY SYS_TYPE "
										  SET RDS=CONN.EXECUTE(SQL)
										  WHILE NOT RDS.EOF
										%>
										<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 21)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
										<%RDS.MOVENEXT
										WEND
										rds.close
										SET RDS=NOTHING
										%>
									</SELECT>
								<%end if%>	
							<%ELSE%>
								<INPUT NAME=shift TYPE=HIDDEN>	
							<%END IF%> 			
						</TD>
						<td>
							<%IF  X_empid<>"" THEN %>
								<%'異動年月<本月
								if  X_calcYM < trim(nowmonth)  and session("Netuser")<>"PELIN"  then  %> 					
									<textarea rows="2" name="memo"  cols="30" onchange="datachg(<%=currentRow-1%>)" wrap="physical"><%=TRIM(tmpRec(CurrentPage, CurrentRow, 23))%></textarea> 					
								<%'異動年月>離職年月	
								elseif  X_calcYM>X_OUTYM and X_outdat<>"" then %>
									<INPUT NAME=memo TYPE=HIDDEN>
								<%else%>
									<textarea rows="2" name="memo"  cols="30" onchange="datachg(<%=currentRow-1%>)" wrap="physical"><%=TRIM(tmpRec(CurrentPage, CurrentRow, 23))%></textarea>
								<%end if%>		
							<%else%>	
								<INPUT NAME=memo TYPE=HIDDEN>
							<%end if%>
						</td> 		
					</TR>
					<%next%>
				<%
				conn.close
				set conn=nothing
				%>	
				</TABLE>
			</td>
		</tr>
		<tr>
			<td align="center">
				<INPUT NAME=JOB TYPE=HIDDEN>
				<INPUT NAME=op TYPE=HIDDEN>
				<INPUT NAME=memo TYPE=HIDDEN>
				<table class="table-borderless table-sm text-secondary txt">
					<tr>
						<td>
						Page:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , Count:<%=RECORDINDB%><BR>
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
						<td><BR>
							<%if UCASE(session("mode"))="W" then%>
								<input type="button" name="send" value="CONFRIM" onclick="GO()" class="btn btn-sm btn-danger">
								<input type="reset" name="send" value="CANCEL" class="btn btn-sm btn-outline-secondary">
							<%end if%>
						</td>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
			
</form>
</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YEBBB0202")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 6) = request("JOB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("memo")(CurrentRow)
	next 
	Session("YEBBB0202") = tmpRec
End Sub
%> 

<script language=javascript>

function datachg(index){
	<%=self%>.op[index].value="upd";
	} 

function gchg(index)
{
	<%=self%>.op[index].value="upd";
	codestr01 = <%=self%>.groupid[index].value.trim();
	open("<%=SELF%>.Back.asp?codestr01=" + codestr01 +"&CurrentPage="+ <%=CurrentPage%> +"&index=" + index + "&func=gchg", "Back");
} 

function oktest(N){
	empid = <%=self%>.f_empid.value;
	ct = <%=self%>.ct.value ;
	
	wt = (window.screen.width )*0.5;
	ht = window.screen.availHeight*0.8;
	tp = (window.screen.width )*0.05;
	lt = (window.screen.availHeight)*0.1;
	if (ct=="VN") 
		open("YEBQ01B.editvn.asp?uid="+ empid +"&empautoid="+ N, "_blank" , "top="+ tp +", left="+ lt +", width="+ wt +",height="+ ht +", scrollbars=yes");
	else
		open("YEBQ01B.edithw.asp?empautoid="+ N , "_blank" , "top="+ tp +", left="+ lt +", width="+ wt +",height="+ ht +", scrollbars=yes");
	 
}

function go(){	
	<%=self%>.action="<%=self%>.updateDB.asp";
	<%=self%>.submit();
} 

</script>

