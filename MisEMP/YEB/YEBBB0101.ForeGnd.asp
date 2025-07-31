<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next 
session.codepage="65001"
SELF = "YEBBB0101"

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

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record
NOWMONTH=CSTR(YEAR(DATE()))&RIGHT("00"&CSTR(MONTH(DATE())),2)
' if RIGHT("00"&CSTR(MONTH(DATE())),2)="12" then 
	' NOWMONTH=CSTR(YEAR(DATE())+1)&"01"
' else	
	' NOWMONTH=CSTR(YEAR(DATE()))&RIGHT("00"&CSTR(MONTH(DATE())+1),2)
' end if 

NOWDAY=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)

if yymm="" then yymm=NOWMONTH

sqlstr = "select * from view_empfile where 1=1 AND whsno like '%"& whsno &"%' and unitno like '%"& unitno &"%'  and groupid like '%"& groupid &"%'  "&_
	"and country like '%"& country &"%'  and zuno like '%"& zuno &"%' and isnull(shift,'') like '%"& shift &"' and  ( empid like '%"& QUERYX &"%' or empnam_VN like '%"& QUERYX &"%'  or empnam_CN like '%"& QUERYX &"%')  "
	if  outemp="Y" then
		sqlstr = sqlstr & " AND isnull(bhdat,'')<>'' "
	elseif 	outemp="N"  then
		sqlstr = sqlstr & " AND isnull(bhdat,'')=''  "
	end if 	
	if EMPID<>"" THEN
		sqlstr = sqlstr & " and EMPID>='"& EMPID &"'  "
	end if 		
	if IOemp="Y" then 
		sqlstr = sqlstr & " AND ( ISNULL(OUTDATE,'')='' OR ISNULL(OUTDATE,'')>='"& NOWDAY &"' )  "
	elseif IOemp="N" then 
		sqlstr = sqlstr & " AND ( ISNULL(OUTDATE,'')<>'' )  "	
	end if 
sqlstr = sqlstr & "order by empid  " 
'response.write sqlstr

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sqlstr, conn, 3, 3
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
			for k=1 to TableRec-1
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = rs("job")
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
				tmpRec(i, j, 23)=""
				tmpRec(i, j, 24)=rs("job")
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
	Session("YEBBB0101") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	gTotalPage = cint(request("gTotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBBB0101")

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
<script language=javascript>

	function f(){
		<%=self%>.yymm.focus();
	}

	function search()
	{		
		<%=self%>.TotalPage.value=0;
		<%=self%>.action="<%=SELF%>.ForeGnd.asp";
		<%=self%>.submit();		
	}
</script>

</head>
<body  topmargin="50" leftmargin="0"  marginwidth="0" marginheight="0"  onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.foreGnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=NowMonth VALUE="<%=NowMonth%>">

<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
<table border=0 width="100%">
	<tr>
		<td align="center">
			<table class="txt"  cellpadding=3 cellspacing=3>
				<TR >
					<TD nowrap align=right >國籍<br>Quoc tich</TD>
					<TD >
						<select name=COUNTRY    onchange="search()" style="width:120px">
							<option value=""></option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							SET RST=NOTHING
							%>
						</SELECT>		
					</TD>
					<TD nowrap align=right>廠別<br>Xuong</TD>
					<TD >
						<select name=WHSNO   onchange="search()"  style="width:120px">
							<option value=""></option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=whsno THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							SET RST=NOTHING
							%>
						</SELECT>
					</TD>					
					<TD nowrap align=right >組/部門<br>Bo phan</TD>
					<TD >
						<select name=GROUPID    onchange="search()" style="width:120px">
							<option value=""></option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=GROUPID THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>					
					<TD nowrap align=right >班別<br>Ca</TD>
					<TD >
						<select name=shift    onchange="search()"  style="width:120px">
							<option value="" <%if shift="" then %> selected<%end if%>></option>
							<option value="ALL" <%if shift="ALL" then %> selected<%end if%>>常日班</option>
							<option value="A" <%if shift="A" then %> selected<%end if%>>A班</option>
							<option value="B" <%if shift="B" then %> selected<%end if%>>B班</option>
						</SELECT>					
					</TD> 
				</TR>
				<TR>	
					<TD nowrap align=right >工號<br>So the</TD>
					<TD nowrap >
						<INPUT type="text"  style="width:100px" NAME=empid1  value="<%=QUERYX%>">			
					</TD>
					<td nowrap align=right >員工統計<br>Thong ke </td>
					<td>
						<select name=IOemp  style="width:120px"> 
							<option value="Y" <%if request("IOemp")="Y" then%>selected<%end if%>>Tai Chuc(在職)</option>
							<option value="" <%if request("IOemp")="" then%>selected<%end if%>>ALL全部</option>
							<option value="N" <%if request("IOemp")="N" then%>selected<%end if%>>Thoi Viec(已離職)</option>
						</select>
					</td>
					<td colspan=4><INPUT TYPE=BUTTON NAME=BTN VALUE="(S)K.Tra查詢" CLASS="btn btn-sm btn-outline-secondary" onclick="search()" ONKEYDOWN="search()"></td>
				</TR>
			</TABLE> 

			<table class="txt"   cellpadding=3 cellspacing=3>
				<tr>
					<td  align="right"><font color=Blue>異動年月(Thang Nam)</font></td>	
					<td><input  type="text" style="width:100px" name=yymm value="<%=yymm%>" onblur="yymmchg()"></td>
					<td>(EX:200701) 只可新增本月以後之晉升資料,其餘異動請至1.2處理</td>	
				</tr>
			</table>
			
			<table id="myTableGrid">	
				<TR class="header">
					<TD nowrap>工號<br>so the</TD>
					<TD nowrap>姓名<br>Ho ten</TD>
					<TD nowrap>到職日期<br>NVX<br>NTV</TD> 					
					<TD nowrap>職等<br>Chuc vu</TD>
					<TD nowrap>異動說明<br>Ly do</TD>		
				</TR>
				<%for CurrentRow = 1 to PageRec
										
				%>
				<TR >
					<TD align=center>
						<a href='javascript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>,<%=currentrow-1%>)'><%=tmpRec(CurrentPage, CurrentRow, 1)%></a> 			
						<INPUT NAME=op TYPE=HIDDEN size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>">
					</TD>
					<TD height=15>
						
						<%IF TRIM(tmpRec(CurrentPage, CurrentRow, 1))<>"" THEN %>
						<a href='javascript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>,<%=currentrow-1%>)'>
							<%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;<font class=txt8VN><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
						</a><br>
						<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 11)%>-<%=tmpRec(CurrentPage, CurrentRow, 13)%></font><BR>
						
						<%end if%>
					</TD>
					<TD align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%>
					<br><font color="red"><%=tmpRec(CurrentPage, CurrentRow, 20)%></font><!--ụ離職-->
					</TD> 		
					<TD align=LEFT><!--職等-->
						<%IF TRIM(tmpRec(CurrentPage, CurrentRow, 1))<>"" and trim(tmpRec(CurrentPage, CurrentRow, 20))="" THEN %>
							<SELECT name="job"  onchange="datachg(<%=currentRow-1%>)">
								<option value="" <%if TRIM(tmpRec(CurrentPage, CurrentRow, 6))="" THEN%>selectd<%end if%>>----</option>
								<%SQL="SELECT * FROM BASICCODE WHERE  FUNC='LEV' ORDER BY SYS_TYPE "
								  SET RDS=CONN.EXECUTE(SQL)
								  WHILE NOT RDS.EOF
								%>
								<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=TRIM(tmpRec(CurrentPage, CurrentRow, 6)) THEN %>SELECTED<%END IF%>><%=RDS("SYS_TYPE")%>-<%=RDS("SYS_VALUE")%></OPTION>
								<%RDS.MOVENEXT
								WEND
								SET RDS=NOTHING
								%>
							</SELECT>
						<%ELSE%>
							<INPUT name="job" TYPE=HIDDEN>	
						<%END IF%>

					</TD>
					<td>
						<%IF TRIM(tmpRec(CurrentPage, CurrentRow, 1))<>"" and trim(tmpRec(CurrentPage, CurrentRow, 20))=""  THEN %>
							<textarea rows="2" name="memo"  cols="25" onchange="datachg(<%=currentRow-1%>)" wrap="physical"><%=TRIM(tmpRec(CurrentPage, CurrentRow, 23))%></textarea>
						<%else%>	
							<INPUT NAME=memo TYPE=HIDDEN>
						<%end if%>
						<INPUT NAME=f_empid TYPE=HIDDEN value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
						<INPUT NAME=ct TYPE=HIDDEN value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>">
					</td> 		
				</TR>
				<%next%>
			</TABLE>
			<INPUT name="job" TYPE=HIDDEN>
			<INPUT NAME=op TYPE=HIDDEN>
			<INPUT NAME=memo TYPE=HIDDEN>
			<INPUT NAME=f_empid TYPE=HIDDEN>
			<INPUT NAME=ct TYPE=HIDDEN>
			<table class="table-borderless table-sm bg-white text-secondary txt">
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

</td></tr></table>
</form>




</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YEBBB0101")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 6) = request("JOB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("memo")(CurrentRow)
	next 
	Session("YEBBB0101") = tmpRec
End Sub
%> 

<script language=javascript>

	function datachg(index){	
		<%=self%>.op[index].value="upd";
		codestr01 = <%=self%>.job[index].value.trim();	
		codestr02 = <%=self%>.memo[index].value.trim();		
		if(codestr02 !=""){
			codestr02 = codestr02.REPLACE("'", "" );	
		} 		
		open("<%=SELF%>.Back.asp?codestr01=" + codestr01 +"&codestr02=" + codestr02 +"&CurrentPage="+ <%=CurrentPage%> +"&index=" + index + "&func=upd", "Back"	);     	     
	} 

	function oktest(N,index){
		//alert(N);
		//alert(index);
		tp=<%=self%>.TotalPage.value;
		cp=<%=self%>.CurrentPage.value;
		rc=<%=self%>.RecordInDB.value ;
		
		empid = <%=self%>.f_empid[index].value;
		ct = <%=self%>.ct[index].value; 
		
		wt = (window.screen.width )*0.5;
		ht = window.screen.availHeight*0.8;
		tp = (window.screen.width )*0.05;
		lt = (window.screen.availHeight)*0.1;
		//alert(ct);
		if(ct=="VN") 
			open("YEBQ01B.editvn.asp?uid="+ empid +"&empautoid="+ N , "_blank" , "top="+ tp +", left="+ lt +", width="+ wt +",height="+ ht +", scrollbars=yes");
		else
			open("YEBQ01B.edithw.asp??uid="+ empid +"&empautoid="+ N , "_blank" , "top="+ tp +", left="+ lt +", width="+ wt +",height="+ ht +", scrollbars=yes");
			
	}

	function yymmchg(){
		if(<%=self%>.yymm.value.trim() !="")
		{ 
			if(<%=self%>.uid.value !="PELIN")
			{ 
				if(<%=self%>.yymm.value.trim() < <%=self%>.NowMonth.value.trim())
				{ 
					alert("只可新增下月以後之資料!!");
					<%=self%>.yymm.focus();
					<%=self%>.yymm.value=<%=self%>.NowMonth.value.trim();
				} 	
			} 	
		} 	
	}

	function go(){
		if(<%=self%>.yymm.value.trim()==""){ 
			alert("請輸入異動年月!!");
			<%=self%>.yymm.focus();			
		}else {	
			<%=self%>.action="<%=self%>.updateDB.asp";
			<%=self%>.submit();
		}
	}
	
</script>

