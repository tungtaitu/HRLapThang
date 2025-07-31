<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
'on error resume next 
session.codepage="65001"
SELF = "YEBBB0102"

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
EMPID = REQUEST("EMPID")
shift = request("shift")
IOemp = request("IOemp") 
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
TableRec = 30    'number of fields per record
NOWMONTH=CSTR(YEAR(DATE()))&RIGHT("00"&CSTR(MONTH(DATE())),2)
'if RIGHT("00"&CSTR(MONTH(DATE())),2)="12" then 
'	NOWMONTH=CSTR(YEAR(DATE())+1)&"01"
'else	
'	NOWMONTH=CSTR(YEAR(DATE()))&RIGHT("00"&CSTR(MONTH(DATE())+1),2)
'end if 

NOWDAY=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)

if yymm="" then yymm=NOWMONTH

sql="select  a.calcyymm, isnull(b.job,'') as CXJob, isnull(b.job,c.job) as JobList, isnull(b.memo,'') as Jobmemo, c.* from  "&_
	"( "&_
	"select  convert(char(6), b.dat, 112) calcyymm,  a.empid  from ( "&_
	"( select  'T' as tmp, empid, indat from empfile  where empid ='"& QUERYX &"' ) a "&_
	"left join ( select  'T' as tmp , *  from  YDBMCALE    ) b  on  b.tmp = a.tmp  and   convert(char(8), b.dat, 112)>=convert(char(8), a.indat, 112)"&_
	")    group by   convert(char(6), b.dat, 112)  , a.empid   "&_
	") a   "&_	
	"left join ( select * from  bempj ) b on b.empid = a.empid and b.yymm = a.calcyymm "&_
	"left join ( select * from view_empfile ) c on c.empid  = a.empid  where left(a.calcyymm,4) <= '"& CSTR(YEAR(DATE())) &"' order by a.calcyymm  desc  "
'response.write sql 	
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 3, 3
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
			F_wstr = rs("wstr")
			F_ustr = RS("ustr")
			F_cstr = RS("cstr")
			F_gstr = RS("gstr")
			F_zstr = RS("zstr") 
			F_shift = RS("sstr") 
			F_BHDAt = RS("bhdat") 
			birDay = rs("BYY")&"/"&right("00"&rs("BMM"),2)&"/"&right("00"&rs("Bdd"),2)
			ages  = rs("ages")
			sex = rs("sex")
			F_indat = rs("nindat")
			F_outdat = rs("outdate")
			country=rs("country")
			
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("empid"))
			tmpRec(i, j, 2) = trim(rs("empnam_cn"))
			tmpRec(i, j, 3) = trim(rs("empnam_vn"))
			tmpRec(i, j, 4) = rs("country")
			tmpRec(i, j, 5) = rs("nindat")
			tmpRec(i, j, 6) = rs("JobList")
			tmpRec(i, j, 7) = rs("whsno")
			tmpRec(i, j, 8) = rs("unitno")
			tmpRec(i, j, 9)	=RS("groupid")
			tmpRec(i, j, 10)=RS("zuno")
			tmpRec(i, j, 11)=RS("wstr")
			tmpRec(i, j, 12)=RS("ustr")
			tmpRec(i, j, 13)=RS("gstr")
			tmpRec(i, j, 14)=RS("zstr")
			tmpRec(i, j, 15)=RS("jstr")
			tmpRec(i, j, 16)=RS("cstr")
			tmpRec(i, j, 17)=RS("autoid")
			IF RS("zuno")="XX" THEN
				tmpRec(i, j, 18)=""
			ELSE
				tmpRec(i, j, 18)=RS("zuno")
			END IF
			tmpRec(i, j, 19)=RS("bhdat")
			tmpRec(i, j, 20)=RS("outdate")
			tmpRec(i, j, 21)=RS("SHIFT")
			tmpRec(i, j, 22)=RS("GTDAT")
			tmpRec(i, j, 23)=replace(rs("jobmemo"),"<BR>", vbCrLf)
			tmpRec(i, j, 24)=rs("CXJob")			
			tmpRec(i, j, 25)=rs("calcyymm")
			if rs("calcyymm") < nowmonth then 
				tmpRec(i, j, 26)="DarkGray"
			elseif rs("calcyymm") = nowmonth then 
				tmpRec(i, j, 26)="Blue"
			else
				tmpRec(i, j, 26)="Firebrick"
			end if 
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
	Session("YEBBB0102") = tmpRec
else 	
	TotalPage = cint(request("TotalPage"))
	gTotalPage = cint(request("gTotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBBB0102")

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
<body  topmargin="40" leftmargin="0"  marginwidth="0" marginheight="0"   >
<form name="<%=self%>" method="post" action="<%=self%>.foreGnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=NowMonth VALUE="<%=NowMonth%>">

<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
<table width="100%" >
	<tr>
		<td>
			<table class="table-borderless table-sm  text-secondary txt">
				<TR>
					<td style="width:20px">&nbsp;</td>
					<TD nowrap align=right height=25  >工號/姓名(So the/Ho Ten): </TD>
					<TD valign="top" colspan=5><%=F_EMPID%>&nbsp;<%=F_name%> (<%=F_cstr%> )
						<INPUT name="f_empid" SIZE=12 CLASS=READONLY VALUE="<%=F_EMPID%>" READONLY  type="hidden" > 
						<INPUT NAME=F_name SIZE=25 CLASS=readonly readonly  VALUE="<%=F_name%>"   type="hidden"  >
						<input type="hidden" name="ct" value="<%=country%>">	
					</TD>							
				</TR> 
				<tr>
					<td style="width:20px">&nbsp;</td>
					<td nowrap align=right>目前單位(Xuong/bo phan) : </td>
					<td  valign="top">(<%=F_wstr%>)&nbsp;&nbsp;&nbsp;<%=F_ustr%>-<%=F_gstr%>-<%=F_zstr%> ( ca:<%=replace(F_shift,"ALL"," ")%> )
						<input name=F_wstr size=10 class='readonly' readonly value="<%=F_wstr%>" type="hidden" >  
						<input name=F_ustr size=10 class='readonly' readonly value="<%=F_ustr%>" type="hidden" > 
						<input name=F_gstr size=5 class='readonly' readonly value="<%=F_gstr%>" type="hidden" > 
						<input name=F_zstr size=5 class='readonly' readonly value="<%=F_zstr%>" type="hidden" > 
						<input name=F_shift size=4 class='readonly' readonly value="<%=F_shift%>"  type="hidden" > 
					</td>
					<TD  nowrap align=right >到職日(NVX) : </TD>
					<TD valign="top"><%=F_indat%>
						<INPUT NAME=F_indat SIZE=12 CLASS=readonly readonly  VALUE="<%=F_indat%>"  type="hidden"  >			
					</TD>
					<TD  nowrap align=right >離職日(NTV) : </TD>
					<TD nowrap><%=F_outdat%>
						<INPUT NAME=F_outdat SIZE=12 CLASS=readonly readonly  VALUE="<%=F_outdat%>"   type="hidden" >			
					</TD>
				</tr>
			</table> 
		</td>
	</tr>
	<tr>
		<td align="center">
			<table id="myTableGrid">				
				<TR class="header"> 		
					<TD>姓名<br>Ho ten</TD>
					<TD>異動年月<br>yymm</TD> 		 		
					<TD>職等<br>Chuc vu</TD>
					<TD>異動說明<br>ghi chu (ly do)</TD>		
				</TR>
				<%for CurrentRow = 1 to PageRec
					 
					X_empid = TRIM(tmpRec(CurrentPage, CurrentRow, 1))  '員工編號
					X_outdat = trim(tmpRec(CurrentPage, CurrentRow, 20)) '離職日
					X_OUTYM = left( trim(tmpRec(CurrentPage, CurrentRow, 20)),4)&mid( trim(tmpRec(CurrentPage, CurrentRow, 20)),6,2) '離職年月
					X_CalcYM = TRIM(tmpRec(CurrentPage, CurrentRow, 25)) '異動年月
					'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
				%>
				<TR> 		
					<TD>
						
						<%IF TRIM(tmpRec(CurrentPage, CurrentRow, 1))<>"" THEN %> 			
						<a href="javascript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)">
							<%=tmpRec(CurrentPage, CurrentRow, 1)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 2)%></font>
						</a>
						<%end if%>
						<INPUT NAME=op TYPE=HIDDEN size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>">
					</TD>
					<TD align=center ><font color="<%=tmpRec(CurrentPage, CurrentRow, 26)%>"><b><%=tmpRec(CurrentPage, CurrentRow, 25)%></b></font></TD> 		
					<TD><!--職等-->
						<%IF  X_empid<>"" THEN %>
							<%'異動年月<本月
							if  X_calcYM < trim(nowmonth) and session("Netuser")<>"PELIN"  then  %>
								<INPUT NAME=job TYPE=HIDDEN value="<%=TRIM(tmpRec(CurrentPage, CurrentRow, 6))%>">
								<SELECT NAME=JOBT  onchange="datachg(<%=currentRow-1%>)" disabled >
									<option value="" <%if TRIM(tmpRec(CurrentPage, CurrentRow, 6))="" THEN%>selectd<%end if%>>----</option>
									<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='LEV' ORDER BY SYS_TYPE "
									  SET RDS=CONN.EXECUTE(SQL)
									  WHILE NOT RDS.EOF
									%>
									<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF trim(RDS("SYS_TYPE"))=TRIM(tmpRec(CurrentPage, CurrentRow, 6)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
									<%RDS.MOVENEXT
									WEND
									SET RDS=NOTHING
									%>
								</SELECT>
							<%'異動年月>=離職年月 and 離職日期<>""	
							elseif  X_calcYM>X_OUTYM and X_outdat<>"" then %>
								<INPUT NAME=job TYPE=HIDDEN>	
							<%else%>
								<SELECT NAME=job  onchange="datachg(<%=currentRow-1%>)">
									<option value="" <%if TRIM(tmpRec(CurrentPage, CurrentRow, 6))="" THEN%>selectd<%end if%>>----</option>
									<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='LEV' ORDER BY SYS_TYPE "
									  SET RDS=CONN.EXECUTE(SQL)
									  WHILE NOT RDS.EOF
									%>
									<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF trim(RDS("SYS_TYPE"))=TRIM(tmpRec(CurrentPage, CurrentRow, 6)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_VALUE")%></OPTION>
									<%RDS.MOVENEXT
									WEND
									rds.close
									SET RDS=NOTHING						
									%>
								</SELECT>
							<%end if%>	
						<%ELSE%>
							<INPUT NAME=job TYPE=HIDDEN>	
						<%END IF%> 			
					</TD>
					<td>
						<%IF  X_empid<>"" THEN %> 				
								<%'異動年月>離職年月	
								if  X_calcYM>X_OUTYM and X_outdat<>"" then %>
									<INPUT NAME=memo TYPE=HIDDEN>
								<%else%>
									<textarea rows="3" name="memo"  cols="30" onchange="datachg(<%=currentRow-1%>)" wrap="physical"><%=TRIM(tmpRec(CurrentPage, CurrentRow, 23))%></textarea>
								<%end if%>		
						<%else%>	
							<INPUT NAME=memo TYPE=HIDDEN>
						<%end if%>
					</td> 		
				</TR>
				<%
					next
					conn.close 
					set conn=nothing
				%>	
			</TABLE>
		</td>
	</tr>
	<tr>
		<td align="center">
			<INPUT NAME=job TYPE=HIDDEN>
			<INPUT NAME=op TYPE=HIDDEN>
			<INPUT NAME=memo TYPE=HIDDEN>
			<TABLE class="table-borderless table-sm bg-white text-secondary txt">
			<tr>
				<td align="CENTER" height=40 width=60%>
				PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%><BR>
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
						<input type="button" name="send" value="CONFRIM" onclick="go()" class="btn btn-sm btn-danger">
						<input type="reset" name="send" value="CANCEL" class="btn btn-sm btn-outline-secondary" >
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
	tmpRec = Session("YEBBB0102")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 6) = request("job")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("memo")(CurrentRow)
	next 
	Session("YEBBB0102") = tmpRec
End Sub
%> 

<script language=javascript>


	function datachg(index){
		
		<%=self%>.op[index].value="upd";
		codestr01 = <%=self%>.job[index].value.trim();	
		codestr02 = <%=self%>.memo[index].value.trim();	 
		
		if(codestr02 !=""){
			codestr02 = codestr02.replace("'", "" );			
		} 
		
		open("<%=SELF%>.Back.asp?codestr01="+ codestr01 +"&codestr02=" + codestr02 +"&CurrentPage="+ <%=CurrentPage%> + "&index=" + index + "&func=upd", "Back");
		
	} 

	function oktest(N){
				
		tp=<%=self%>.TotalPage.value;		
		cp=<%=self%>.CurrentPage.value;		
		rc=<%=self%>.RecordInDB.value;		
		empid = <%=self%>.f_empid.value;		
		ct = <%=self%>.ct.value ;
		
		wt = (window.screen.width )*0.5;
		ht = window.screen.availHeight*0.8;
		tp = (window.screen.width )*0.05;
		lt = (window.screen.availHeight)*0.1;
				
		if(ct=="VN") 
			open("YEBQ01B.editvn.asp?uid="+ empid +"&empautoid="+ N, "_blank" , "top="+ tp +", left="+ lt +", width="+ wt +",height="+ ht +", scrollbars=yes");
		else
			open("YEBQ01B.edithw.asp?empautoid="+ N , "_blank" , "top="+ tp +", left="+ lt +", width="+ wt +",height="+ ht +", scrollbars=yes");	
		
	}

	function go(){	
		<%=self%>.action="<%=self%>.updateDB.asp";
		<%=self%>.submit();
	} 

</script>

