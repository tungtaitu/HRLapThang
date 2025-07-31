<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "empfilesalary"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   

whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
zuno = trim(request("zuno"))
job = trim(request("job"))
QUERYX = trim(request("QUERYX")) 

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record  


sql=" select b.sys_value as whsnodesc, c.sys_value as unitdesc , d.sys_value as groupdesc , "&_
	"e.sys_value as zunodesc, f.sys_value as jobdesc, g.sys_value as countrydesc, "&_
	"convert(char(10),isnull(indat,''),111) as date1, h.code as cvcode , h.bonus as cv_bonus , "&_
	"i.bonus as bb_bonus, a.*  "&_
	"from     "&_
	"( select * from  empfile  ) a  "&_
	"left join ( select * from basicCode where func ='whsno' ) b on b.sys_type=a.whsno  "&_
	"left join ( select * from basicCode where func  ='unit' ) c on c.sys_type = a.unitno  "&_
	"left join ( select * from basicCode where func  ='groupid' ) d on d.sys_type = a.groupid  "&_
	"left join ( select * from basicCode where func  ='zuno' ) e on e.sys_type = a.zuno  "&_
	"left join ( select * from basicCode where func  ='lev' ) f on f.sys_type = a.job "&_
	"left join ( select * from basicCode where func  ='country' ) g on g.sys_type = a.country  "&_ 
	"left join ( select * from empsalarybasic where func='BB' and yymm='200603') h on h.job = a.job "&_ 
	"left join ( select * from empsalarybasic where func='AA' and yymm='200603' ) i on i.code = a.bb "&_ 
	"where ISNULL(A.STATUS,'')<>'D' AND ( isnull(a.outdat,'')='' or convert(char(10),a.outdat,111)> convert(char(10), getdate(), 111) ) and whsno like '%"& whsno &"%' and unitno like '%"& unitno &"%'  and groupid like '%"& groupid &"%'  "&_
	"and zuno like '%"& zuno &"%' and A.job like '%"& job &"%' and empid like '%"& QUERYX &"%'  "
if trim(request("empidstr"))="" then 	
	sql=sql&"order by empid" 
else 
	sql=sql &"and empid in ( " & left(request("empidstr"),len(trim(request("empidstr")))-1) & ") "
	sql=sql &"order by empid " 
end  if 	
response.write sql 
response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
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
			for k=1 to TableRec-1
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("date1")
				tmpRec(i, j, 6) = rs("job")				
				tmpRec(i, j, 7) = rs("whsno")	 
				tmpRec(i, j, 8) = rs("unitno")	 
				tmpRec(i, j, 9)	=RS("groupid") 
				tmpRec(i, j, 10)=RS("zuno") 				
				tmpRec(i, j, 11)=RS("whsnodesc") 	
				tmpRec(i, j, 12)=RS("unitdesc") 	
				tmpRec(i, j, 13)=RS("groupdesc") 	
				tmpRec(i, j, 14)=RS("zunodesc") 	
				tmpRec(i, j, 15)=RS("jobdesc") 	
				tmpRec(i, j, 16)=RS("countrydesc") 	
				tmpRec(i, j, 17)=RS("autoid") 	
				IF RS("zuno")="XX" THEN 
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("zuno")
				END IF 
				tmpRec(i, j, 19)=RS("BB")
				tmpRec(i, j, 20)=RS("BB_bonus")
				tmpRec(i, j, 21)=RS("CVcode")
				tmpRec(i, j, 22)=RS("CV_bonus")
				tmpRec(i, j, 23)=RS("PHU")		
				tmpRec(i, j, 24)=RS("NN")
				tmpRec(i, j, 25)=RS("KT")
				tmpRec(i, j, 26)=RS("MT")
				tmpRec(i, j, 27)=RS("TTKH")
						
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
	Session("empfilesalary") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empfilesalary")
	
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
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	<%=self%>.QUERYX.focus()	
end function   

function chgdata()
	<%=self%>.action="empfile.salary.asp?totalpage=0"
	<%=self%>.submit
end function 

-->
</SCRIPT>  
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="empfile.salary.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD >
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	人事薪資系統( 員工基本資料-薪資管理 ) </TD>	
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<TABLE WIDTH=460 CLASS=FONT9 BORDER=0>   
	<TR height=25 >
		<TD nowrap align=right>廠別</TD>
		<TD > 
			<select name=WHSNO  class=font9 onchange="chgdata()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=whsno THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD> 
		<TD   nowrap align=right>處/所</TD>
		<TD > 
			<select name=unitno  class=font9 onchange="chgdata()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='unit' and sys_type<>'AAA' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=unitno THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>	
		<TD nowrap align=right >組/部門</TD>
		<TD >
			<select name=GROUPID  class=font9  onchange="chgdata()">
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
		<TD nowrap align=right >單位</TD>
		<TD >
			<select name=zuno  class=font9  onchange="chgdata()" >
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='ZUNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=ZUNO THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>			 				 
			</SELECT>
			<%SET RST=NOTHING %>			
		</TD>
	</TR>	
	<TR>
		<TD nowrap align=right>職等</TD>
		<TD COLSPAN=3> 
			<select name=JOB  class=font9 onchange="chgdata()" >
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>"  <%IF RST("SYS_TYPE")=JOB THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD> 
		<TD nowrap align=right >關鍵字</TD> 
		<TD COLSPAN=3>
			<INPUT NAME=QUERYX SIZE=18 CLASS=INPUTBOX value="<%=QUERYX%>">
			<INPUT TYPE=BUTTON NAME=BTN VALUE="查詢" CLASS=BUTTON onclick="chgdata()" >
		</TD>		
	</TR>
</TABLE>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>	 	
<!--------------------------------------------------------------------> 		
<TABLE WIDTH=550 CLASS="FONT9" BORDER=0 cellspacing="0" cellpadding="1" > 	
	<TR BGCOLOR="LightGrey" HEIGHT=25   >
 		<TD width=55 nowrap align=center ROWSPAN=2>工號</TD> 		
 		<TD width=100 nowrap align=LEFT ROWSPAN=2 >姓名</TD> 
 		<td nowrap align=center>到職日期</td>		
 		<TD width=60 nowrap align=center>薪資代碼</TD>
 		<TD width=60 nowrap align=center>基本薪資</TD>
 		<TD width=60 nowrap align=center>職專</TD> 			
 		<TD width=60 nowrap align=center>職專加給</TD>
 	</TR>
 	<tr BGCOLOR="LightGrey"  HEIGHT=25 > 		
 		<TD width=60 nowrap align=center>特加(Y)</TD>
 		<td nowrap align=center>語言加給</td>
 		<td nowrap align=center>技術加給</td>
 		<td nowrap align=center>環境加給</td>
 		<td nowrap align=center>其他加給</td>
 	</tr> 
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR="#DFEFFF"
		END IF 	 
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
	%>
	<TR BGCOLOR=<%=WKCOLOR%> > 		
 		<TD align=center ROWSPAN=2>
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 1)%>
 			</a>
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
 		</TD> 		
 		<TD ROWSPAN=2>
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 2)%><BR>
 				<font class=txt8VN><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
 			</a>
 		</TD>
 		<TD align=center><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></TD>
 		<TD ALIGN=RIGHT >
 			<%'=tmpRec(CurrentPage, CurrentRow, 19)%>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			
		 		<select name=BBCODE  class="txt8" style="width:60" onchange="bbcodechg(<%=currentrow-1%>)">				
					<%SQL="SELECT * FROM empsalarybasic WHERE FUNC='AA'  ORDER BY BONUS "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("CODE")%>" <%IF RST("CODE")=trim(tmpRec(CurrentPage, CurrentRow, 19)) THEN %> SELECTED <%END IF%> ><%=RST("CODE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>
				<input type=hidden name=BBCODE >	
			<%end if %>
 		</TD>
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			
	 			<%'=tmpRec(CurrentPage, CurrentRow, 20)%>
	 			<INPUT NAME=BB CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY >	 			
	 		<%else%>
				<input type=hidden name=BB >	
			<%end if %>	
 		</TD>
 		<TD ALIGN=RIGHT ><!--職等-->
 			<%'=tmpRec(CurrentPage, CurrentRow, 6)%>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<select name=F1_JOB  class="txt8" style="width:60" ONCHANGE="JOBCHG(<%=CURRENTROW-1%>)">				
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=trim(tmpRec(CurrentPage, CurrentRow, 6)) THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>
				<input type=hidden name=F1_JOB >	
			<%end if %>
 		</TD>
 		 <TD  ALIGN=RIGHT>
 		 	<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>				
 		 		<%'=tmpRec(CurrentPage, CurrentRow, 22)%>
 		 		<INPUT NAME=CV CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" STYLE="TEXT-ALIGN:right;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY >
 		 		<input type=hidden name=CVCODE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" SIZE=3>
 		 	<%else%>
				<input type=hidden name=CV >	
				<input type=hidden name=CVCODE >
			<%end if %>	
 		 </TD>
 		 
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> >
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<%'=tmpRec(CurrentPage, CurrentRow, 23)%>
	 			<INPUT NAME=PHU CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)">
	 		<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=PHU	>
			<%END IF%>	
 		</TD> 		
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<%'=tmpRec(CurrentPage, CurrentRow, 24)%>
 				<INPUT NAME=NN CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=NN >
			<%END IF%>		
 		</TD>
 		<TD  ALIGN=RIGHT>
	 		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<%'=tmpRec(CurrentPage, CurrentRow, 25)%>
	 			<INPUT NAME=KT CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KT >
	 		<%END IF%>			
 		</TD>
 		<TD  ALIGN=RIGHT>
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 		
	 			<%'=tmpRec(CurrentPage, CurrentRow, 26)%>
	 			<INPUT NAME=MT CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=MT >
	 		<%END IF%>			
 		</TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<%'=tmpRec(CurrentPage, CurrentRow, 27)%> 			
	 			<INPUT NAME=TTKH CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TTKH >
	 		<%END IF%>			
 		</TD>
	</TR>
	
	<%next%>
</TABLE>	
<input type=hidden name=empid>
<input type=hidden name=BBCODE>
<input type=hidden name=BB>
<input type=hidden name=F1_JOB>
<input type=hidden name=CV>
<input type=hidden name=CVCODE>
<input type=hidden name=PHU>
<input type=hidden name=NN>
<input type=hidden name=KT>
<input type=hidden name=MT>
<input type=hidden name=TTKH>


<TABLE border=0 width=500 class=font9 >
<tr>
    <td align="CENTER" height=40 WIDTH=75%>    
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
	<FONT CLASS=TXT8>&nbsp;&nbsp;PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%></FONT>
	</TD>	
	<TD WIDTH=25% ALIGN=RIGHT>		
		<input type="BUTTON" name="send" value="確　認" class=button ONCLICK="GO()">
		<input type="BUTTON" name="send" value="取　消" class=button onclick="clr()">
	</TD>
</TR>
	
</TABLE> 
</form>
  



</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("empfilesalary")
	for CurrentRow = 1 to PageRec
		'tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		 
		tmpRec(CurrentPage, CurrentRow, 6) = request("F1_JOB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 19) = request("BBCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 20) = request("BB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 21) = request("CVCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 22) = request("CV")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("PHU")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 24) = request("NN")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 25) = request("KT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 26) = request("MT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 27) = request("TTKH")(CurrentRow)
		
	next 
	Session("empfilesalary") = tmpRec
	
End Sub
%> 

<script language=vbscript>
function BACKMAIN() 	
	open "../main.asp" , "_self"
end function   

function clr()
	open "empfile.salary.asp" , "_self"
end function 

function go()
	<%=self%>.action="empsalary.upd.asp"  
	<%=self%>.submit()
end function 

function oktest(N)
	tp=<%=self%>.totalpage.value 
	cp=<%=self%>.CurrentPage.value 
	rc=<%=self%>.RecordInDB.value 
	open "empfile.show.asp?empautoid="& N , "_blank" , "top=10, left=10, width=550, scrollbars=yes" 
end function 

FUNCTION BBCODECHG(INDEX)
	codestr=<%=self%>.bbcode(index).value 
	open "empsalary.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&code=" &	codestr , "Back"
		 
	'DATACHG(INDEX)	 
	 
	'PARENT.BEST.COLS="70%,30%"	 	
END FUNCTION 

FUNCTION JOBCHG(INDEX)
	codestr=<%=self%>.F1_JOB(index).value 
	open "empsalary.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&code=" &	codestr , "Back" 
	'PARENT.BEST.COLS="70%,30%"	 	 
	'DATACHG(INDEX)	  
END FUNCTION 

FUNCTION DATACHG(INDEX) 	 	
	if isnumeric(<%=SELF%>.PHU(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.phu(index).focus()
		<%=self%>.phu(index).value=0
		<%=self%>.phu(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.NN(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.NN(index).value=0 		
		<%=self%>.NN(index).focus()
		<%=self%>.NN(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.KT(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.KT(index).value=0 		
		<%=self%>.KT(index).focus()
		<%=self%>.KT(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.MT(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.MT(index).value=0 		
		<%=self%>.MT(index).focus()
		<%=self%>.MT(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.TTKH(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.TTKH(index).value=0 		
		<%=self%>.TTKH(index).focus()
		<%=self%>.TTKH(index).select()
		exit FUNCTION 
	end if 	 
	
	CODESTR01 = <%=SELF%>.PHU(INDEX).VALUE
	CODESTR02 = <%=SELF%>.NN(INDEX).VALUE
	CODESTR03 = <%=SELF%>.KT(INDEX).VALUE
	CODESTR04 = <%=SELF%>.MT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.TTKH(INDEX).VALUE
	'ALERT CODESTR02
	'ALERT CODESTR03
	
	
	open "empsalary.back.asp?ftype=CDATACHG&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 , "Back"  
		 
	'PARENT.BEST.COLS="70%,30%"	 
	
END FUNCTION  
	
</script>

