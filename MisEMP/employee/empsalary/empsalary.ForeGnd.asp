<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "empsalaryForeGnd"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
YYMM=REQUEST("YYMM")
whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
job = trim(request("job"))
QUERYX = trim(request("empid1"))  
outemp = request("outemp")
lastym = left(yymm,4) &  right("00" & cstr(right(yymm,2)-1) ,2 )
if right(yymm,2)="01"  then 
	lastym = left(yymm,4)-1 &"12" 
end if 	

calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"   
'下個月
if right(yymm,2)="12" then 
	ccdt = cstr(left(YYMM,4)+1)&"/01/01" 
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"  
end if 	 
'response.write ccdt  
 
 '一個月有幾天 
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS) 
      

'本月假日天數 (星期日)
SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' AND  DATEPART( DW,DAT ) ='1'  " 
Set rsTT = Server.CreateObject("ADODB.Recordset")   
RSTT.OPEN SQL, CONN, 3, 3 
IF NOT RSTT.EOF THEN 
	HHCNT = CDBL(RSTT.RECORDCOUNT)
ELSE
	HHCNT = 0 
END IF 
SET RSTT=NOTHING   

'RESPONSE.WRITE HHCNT &"<br>" 
'RESPONSE.END  
'本月應記薪天數 
MMDAYS = CDBL(days)-CDBL(HHCNT) 
'RESPONSE.WRITE  MMDAYS 
'RESPONSE.END 
'----------------------------------------------------------------------------------------


gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 70    'number of fields per record  

sql="select * from view_employee "&_	
	"where CONVERT(CHAR(10), indat, 111)< '"& ccdt &"' and ( isnull(outdat,'')='' or outdat>'"& calcdt &"' )  "&_
	"and whsno like '%"& whsno &"%' and unitno like '%"& unitno &"%' and groupid like '%"& groupid &"%'  "&_
	"and COUNTRY like '%"& COUNTRY  &"%' and job like '%"& job &"%' and empid like '%"& QUERYX &"%' " 
	if outemp="D" then  
		sql=sql&" and ( isnull(outdat,'')<>'' and  outdat>'"& calcdt &"' )  " 
	end if
	
sql=sql&"order by empid   "
	
 
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
				tmpRec(i, j, 20)=RS("BB_bonus")  '基本薪資
				tmpRec(i, j, 21)=RS("CVcode")
				tmpRec(i, j, 22)=RS("CV_bonus")  '職務加給
				tmpRec(i, j, 23)=RS("PHU")		'Y獎金 
				tmpRec(i, j, 24)=RS("NN")  '語言加給
				tmpRec(i, j, 25)=RS("KT") '技術加給
				tmpRec(i, j, 26)=RS("MT") '環境加給
				tmpRec(i, j, 27)=RS("TTKH")  '其他加給
				tmpRec(i, j, 28)=RS("BHDAT") '買保險日期
				tmpRec(i, j, 29)=RS("GTDAT") '工團日期
				tmpRec(i, j, 30)=RS("OUTDAT") '離職日期 				
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
	Session("empsalaryForeGnd") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empsalaryForeGnd")
	
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
	'<%=self%>.PHU(0).focus()	
	'<%=self%>.PHU(0).SELECT()
end function   

function chgdata()
	<%=self%>.action="empfile.salary.asp?totalpage=0"
	<%=self%>.submit
end function 
-->
</SCRIPT>  
</head>   
<body   topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()" bgproperties="fixed"  >
<form name="<%=self%>" method="post" action="empfile.salary.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>"> 	
<INPUT TYPE=hidden NAME=MMDAYS VALUE="<%=MMDAYS%>"><!--本月工作天數--> 
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	人事薪資系統( 員工薪資管理 )　
	計薪年月：<%=YYMM%></TD>	
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<!--TABLE WIDTH=460 CLASS=FONT9 BORDER=0>   
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
		<TD nowrap align=right >國籍</TD>
		<TD >
			<select name=COUNTRY  class=font9  onchange="chgdata()" >
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=COUNTRY THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>				 
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
		<TD nowrap align=right>職等</TD>
		<TD COLSPAN=3> 
			<select name=JOB  class=font9 onchange="chgdata()" STYLE='WIDTH:100'>
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
		</TD --> 
		<!--TD nowrap align=right >工號</TD> 
		<TD COLSPAN=3>
			<INPUT NAME=QUERYX SIZE=8 CLASS=INPUTBOX value="<%=QUERYX%>">
			<INPUT TYPE=BUTTON NAME=BTN VALUE="查詢" CLASS=BUTTON onclick="chgdata()" >
		</TD--> 
	<!--/TR>	
</TABLE>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500 -->	 	
<!--------------------------------------------------------------------> 		
<TABLE  CLASS="FONT9" BORDER=0 cellspacing="0" cellpadding="1" > 	
	<TR HEIGHT=25 BGCOLOR="LightGrey"   >
 		<TD ROWSPAN=2 >項次</TD>
 		<TD align=center>工號</TD> 		
 		<TD COLSPAN=4  >員工姓名(中,英,越)</TD> 
 		<td  align=center>時薪</td>
 		<td align=center>到職日期</td>
 		<td align=center>離職日期</td>
 		<TD align=center>工作天數</TD>
 		<TD align=center>上月補款</TD>
 		<TD align=center>績效獎金</TD>
 		<TD align=center>總加班費</TD>
 		<td align=center>(-)工團費</td>
 		<td align=center>(-)其他</td> 	
 		<td align=center>(-)扣時假</td>
 		<TD align=center>應發工資</TD> 		
 		<TD COLSPAN=4 ALIGN=CENTER bgcolor="#ccff99">加班(H)</TD>
 		<TD bgcolor="#ffcc99"></TD>
 		<TD bgcolor="#ffcc99"></TD>	
 		<TD COLSPAN=2 ALIGN=CENTER bgcolor="#ffcccc">請假(H)</TD> 		
 	</TR>
 	<tr BGCOLOR="LightGrey"  HEIGHT=25 > 	
 		<TD align=center>薪資代碼</TD>
 		<TD align=center>基本薪資</TD>
 		<TD align=center>職專</TD> 			
 		<TD align=center>職專加給</TD>	
 		<TD align=center>獎金(Y)</TD>
 		<td align=center>語言加給</td>
 		<td align=center>技術加給</td>
 		<td align=center>環境加給</td>
 		<td align=center>其他加給</td>
 		<td align=center>全勤獎金</td>
 		<td align=center>其他收入</td>  	
 		<TD align=center>應領薪資</TD>	
 		<td align=center>(-)保險費</td>
 		<td align=center>(-)伙食費</td>
 		<TD align=center>離職補助</TD>
 		<TD ALIGN=CENTER >實領工資</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">平日</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">休息</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">假日</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">夜班</TD>  
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>曠職</font></TD>
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>忘遲</font></TD>		
 		<TD ALIGN=CENTER bgcolor="#ffcccc" >事假</TD>
 		<TD ALIGN=CENTER bgcolor="#ffcccc" >病假</TD> 		
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
		<TD ROWSPAN=2 ALIGN=CENTER >
		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*10)%><%END IF %>
		</TD>
 		<TD  >
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 1)%>
 			</a>
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
 			<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
 		</TD> 		
 		<TD COLSPAN=4>
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 2)%>
 				<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
 			</a>
 		</TD>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
 				<INPUT NAME=HHMOENY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 38)%>" CLASS='INPUTBOX8' SIZE=7 STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc;BACKGROUND-COLOR:LIGHTYELLOW" >  
 			<%ELSE%>	
 				<INPUT NAME=HHMOENY TYPE=HIDDEN> 
 			<%END IF %>	 				
 		</TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 30),8)%></FONT></TD>
 		<TD > 		 		
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=WORKDAYS CLASS='INPUTBOX8' READONLY  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 59)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=WORKDAYS >
	 		<%END IF%>	
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=TBTR CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 33)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 上月補款" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TBTR >
	 		<%END IF%>	
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JX CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 58)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 績效獎金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JX >
	 		<%END IF%>	
 		</TD>
 		<TD > 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTJB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 49)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 總加班費">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTJB >
	 		<%END IF%>	  		
 		</TD>
 		<!--TD ALIGN=CENTER ><FONT CLASS=TXT8 COLOR="SeaGreen"><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 28),8)%></FONT></TD> 		
 		<TD ALIGN=CENTER ><FONT CLASS=TXT8><%IF (tmpRec(CurrentPage, CurrentRow, 29))<>"" THEN%>Y<%END IF%></FONT></TD-->
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=GT CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 36)%>" STYLE="TEXT-ALIGN:RIGHT"  onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工團費" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=GT >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=QITA CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 37)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣除其他" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QITA >
	 		<%END IF%>	
 		</TD>  				 		  		
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTKJ CLASS='INPUTBOX8' SIZE=9 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 50)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣時假">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTKJ >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=TOTMONEY CLASS='INPUTBOX8' SIZE=10 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 39),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 應發工資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTMONEY  >
	 		<%END IF%>	
 		</TD> 
 		<TD COLSPAN=4 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" onclick="view1(<%=currentrow-1%>)"><%=tmpRec(CurrentPage, CurrentRow, 1)%> 出勤紀錄</div></u></font>
 			<%END IF %>	
 		</TD> 
 		<TD ></TD>
 		<TD ></TD>
 		<TD COLSPAN=2 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" >看請假紀錄</div></u></font>
 			<%END IF %>		
 		</TD>  		
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> >
 		<TD ALIGN=RIGHT > 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			
		 		<select name=BBCODE  class="txt8" style="width:60" onchange="bbcodechg(<%=currentrow-1%>)">				
					<%SQL="SELECT * FROM empsalarybasic WHERE FUNC='AA'  ORDER BY CODE "
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
	 			<INPUT NAME=BB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">	 			
	 		<%else%>
				<input type=hidden name=BB >	
			<%end if %>	
 		</TD>
 		<TD ALIGN=RIGHT ><!--職等--> 			
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
 		 		<INPUT NAME=CV CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" STYLE="TEXT-ALIGN:right;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 職務加給" >
 		 		<input type=hidden name=CVCODE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" SIZE=3>
 		 	<%else%>
				<input type=hidden name=CV >	
				<input type=hidden name=CVCODE >
			<%end if %>	
 		 </TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=PHU CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助獎金(Y)" >
	 		<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=PHU	>
			<%END IF%>	
 		</TD> 		
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME=NN CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 語言加給" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=NN >
			<%END IF%>		
 		</TD>
 		<TD  ALIGN=RIGHT>
	 		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=KT CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 技術加給" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KT >
	 		<%END IF%>			
 		</TD>
 		<TD  ALIGN=RIGHT>
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 
	 			<INPUT NAME=MT CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 環境加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=MT >
	 		<%END IF%>			
 		</TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=TTKH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TTKH >
	 		<%END IF%>			
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=QC CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 全勤">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QC >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TNKH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他收入">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TNKH >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 60)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:#cc0000"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 應領薪資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTM >
	 		<%END IF%>	
 		</TD>  
 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=BH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 保險費"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BH >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=HS CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 35)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 伙食費" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=HS >
	 		<%END IF%>	
 		</TD>
 		
 		
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=LZBZJ CLASS='INPUTBOX8' SIZE=9 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 62)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=LZBZJ >
	 		<%END IF%>
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=RELTOTMONEY CLASS='INPUTBOX8' VALUE="<%=formatnumber( tmpRec(CurrentPage, CurrentRow, 47),0)%>" SIZE=10  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;;color:#cc0000" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實領工資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=RELTOTMONEY  >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=H1 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 40)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H1 >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=H2 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 41)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H2 >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=H3 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 42)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H3 >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=B3 CLASS='INPUTBOX8'  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 43)%>"  SIZE=4  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=B3 >
	 		<%END IF%>	
 		</TD> 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=KZHOUR CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 44)%>"  SIZE=4  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:Brown" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KZHOUR >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=Forget CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 48)%>"  SIZE=4  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:Brown" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=Forget >
	 		<%END IF%>	
 		</TD> 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=JIAA CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 45)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:ForestGreen" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JIAA >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=JIAB CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 46)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:ForestGreen" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JIAB >
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
	tmpRec = Session("empsalaryForeGnd")
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
	Session("empsalaryForeGnd") = tmpRec
	
End Sub
%> 

<script language=vbscript>
function BACKMAIN() 	
	open "../main.asp" , "_self"
end function   

function clr()
	open "empsalary.fore.asp" , "_self"
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
	daystr=<%=self%>.MMDAYS.value 	
	open "empsalary.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" & daystr & "&code=" &	codestr , "Back" 		 
	'DATACHG(INDEX)	  
	 
	'PARENT.BEST.COLS="70%,30%"	 	
END FUNCTION 

FUNCTION JOBCHG(INDEX)
	codestr=<%=self%>.F1_JOB(index).value 
	daystr=<%=self%>.MMDAYS.value 
	open "empsalary.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" &daystr & "&code=" &	codestr , "Back"
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
	
	if isnumeric(<%=SELF%>.TNKH(INDEX).VALUE)=false then  '其他收入
		alert "請輸入數字!!"
		<%=self%>.TNKH(index).value=0 		
		<%=self%>.TNKH(index).focus()
		<%=self%>.TNKH(index).select()
		exit FUNCTION 
	end if 	 
	
	if isnumeric(<%=SELF%>.HS(INDEX).VALUE)=false then  '伙食費(-)
		alert "請輸入數字!!"
		<%=self%>.HS(index).value=0 		
		<%=self%>.HS(index).focus()
		<%=self%>.HS(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.QITA(INDEX).VALUE)=false then  '其他扣除額(-)
		alert "請輸入數字!!"
		<%=self%>.QITA(index).value=0 		
		<%=self%>.QITA(index).focus()
		<%=self%>.QITA(index).select()
		exit FUNCTION 
	end if    
	
	if isnumeric(<%=SELF%>.JX(INDEX).VALUE)=false then  '其他扣除額(-)
		alert "請輸入數字!!"
		<%=self%>.JX(index).value=0 		
		<%=self%>.JX(index).focus()
		<%=self%>.JX(index).select()
		exit FUNCTION 
	end if  
	TTM = ( cdbl(<%=self%>.bb(index).value) + cdbl(<%=self%>.CV(index).value) + cdbl(<%=self%>.PHU(index).value) ) 
	if TTM mod (26*8)<>0 then 
		TTMH = FIX (CDBL(TTM)/26/8 ) +1   '時薪
	else
		TTMH = FIX (CDBL(TTM)/26/8 )    '時薪
	end if 
	'alert  TTMH 
	'<%=self%>.HHMOENY(index).value = TTMH 
	
	CODESTR01 = <%=SELF%>.PHU(INDEX).VALUE
	CODESTR02 = <%=SELF%>.NN(INDEX).VALUE
	CODESTR03 = <%=SELF%>.KT(INDEX).VALUE
	CODESTR04 = <%=SELF%>.MT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.TTKH(INDEX).VALUE
	CODESTR06 = <%=SELF%>.TNKH(INDEX).VALUE
	CODESTR07 = <%=SELF%>.HS(INDEX).VALUE
	CODESTR08 = <%=SELF%>.QITA(INDEX).VALUE
	CODESTR09 = <%=SELF%>.JX(INDEX).VALUE 
	CODESTR10 = <%=SELF%>.BH(INDEX).VALUE 
	CODESTR11 = <%=SELF%>.GT(INDEX).VALUE 
	daystr=<%=self%>.MMDAYS.value  
	'ALERT CODESTR02
	'ALERT CODESTR03
	
	open "empsalary.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 &_
		 "&CODESTR06="& CODESTR06 &_
		 "&CODESTR07="& CODESTR07 &_	
		 "&CODESTR08="& CODESTR08 &_
		 "&CODESTR09="& CODESTR09 &_
		 "&CODESTR10="& CODESTR10 &_
		 "&CODESTR11="& CODESTR11 &"&days=" & daystr , "Back"  
		 
	'PARENT.BEST.COLS="70%,30%"	 
	
END FUNCTION  

function view1(index) 
	 
	yymmstr = <%=self%>.yymm.value 
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value 
	idstr=  <%=self%>.empautoid(index).value 
	open "empworkb.fore.asp?yymm=" & yymmstr &"&EMPID=" & empidstr &"&empautoid=" & idstr , "_blank" , "top=10, left=10, scrollbars=yes" 
end function 
	
</script>

