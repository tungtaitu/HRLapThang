<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "YECE0101"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
YYMM=REQUEST("YYMM")
whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
F_country = trim(request("F_country"))
job = trim(request("job"))
QUERYX = trim(request("empid1"))  
outemp = request("outemp")
lastym = left(yymm,4) &  right("00" & cstr(right(yymm,2)-1) ,2 )
nowmonth = left(year(date()),4) &  right("00" & cstr(month(date())) ,2 ) 
 

calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"   
'下個月
if right(yymm,2)="12" then 
	ccdt = cstr(left(YYMM,4)+1)&"/01/01" 
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"  
end if 	 
'response.write ccdt  
 


if right(yymm,2)="01"  then 
	lastym = left(yymm,4)-1 &"12" 
end if 	
 
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
TableRec = 37    'number of fields per record  

 
sql="select datediff( m, indat, '"& calcdt &"'  ) as nz,  isnull(a.job, b.job) as bjob, a.whsno as bwhsno, "&_
	"a.dm as Bdm, case when isnull(a.bb,0)<>0 and isnull(a.bb,0)<b.bb then b.bb else isnull(a.bb, 0) end as B_bb, "&_
	"isnull(a.cv,b.cv ) as B_CV, isnull(a.phu,b.phu) as B_PHU , isnull(a.nn,0) as B_NN, "&_ 
	"isnull(a.kt,b.kt) as B_KT, isnull(a.mt,b.mt) as B_MT, isnull(a.ttkh,b.ttkh) as B_TTKH, "&_
	"isnull(a.qc,b.qc) B_QC, a.memo as B_memo , b.* from "&_
	"( select * from view_empfile  ) b "&_
	"left join ( select * from bemps  where  yymm='"& yymm  &"' ) a  on a.empid = b.empid  "&_ 
	"where CONVERT(CHAR(10), b.indat, 111)< '"& ccdt &"' and ( isnull(outdat,'')='' or outdat>'"& calcdt &"' )  "&_
	"and b.whsno like '"& whsno &"%' and b.unitno like '%"& unitno &"%' and b.groupid like '"& groupid &"%'  "&_
	"and b.COUNTRY like '"& F_country  &"%' and b.job like '"& job &"%' and b.empid like '%"& QUERYX &"%' " 
	if outemp="D" then  
		sql=sql&" and ( isnull(outdat,'')<>'' and  outdat>'"& calcdt &"' )  " 
	elseif len(outemp)=6 then
		sql=sql&" and  convert(char(6),b.indat,112)='"& outemp &"'    " 
	end if  
sql=sql&"order by b.empid"
 
response.write trim(sql )
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
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat") 
				'if yymm<>mowmonth then 					
					tmpRec(i, j, 6) = rs("bjob")				
				'else
				'	tmpRec(i, j, 6) = rs("job")
				'end if 	
				if yymm<>mowmonth then 
					tmpRec(i, j, 7) = rs("bwhsno")	 
				else
					tmpRec(i, j, 7) = rs("whsno")	 
				end if	
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
				tmpRec(i, j, 19)=RS("BB")
				if RS("B_BB")="0" then 
					tmpRec(i, j, 20)=cdbl(rs("bb"))
				else
					tmpRec(i, j, 20)=cdbl(RS("B_BB"))  '基本薪資
				end if 	
				
				tmpRec(i, j, 21)=""  'RS("CVcode")
				if rs("country")="VN" then 
					tmpRec(i, j, 22)=cdbl(RS("B_CV"))  '職務加給
				else
					tmpRec(i, j, 22)=cdbl(rs("B_CV"))
				end if 	
				tmpRec(i, j, 23)=cdbl(RS("B_PHU"))		'Y獎金 (陸幹為其他加給)
				tmpRec(i, j, 24)=cdbl(RS("B_NN"))  '語言加給
				tmpRec(i, j, 25)=cdbl(RS("B_KT")) '技術加給
				tmpRec(i, j, 26)=cdbl(RS("B_MT")) '環境加給(陸幹為年資加給)
				tmpRec(i, j, 27)=cdbl(RS("B_TTKH"))  '其他加給(陸幹為補助醫療)
				tmpRec(i, j, 28)=RS("BHDAT") '買保險日期
				tmpRec(i, j, 29)=RS("GTDAT") '工團日期
				tmpRec(i, j, 30)=RS("OUTDATE") '離職日期 		 
				TOTY=  CDBL( ( CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23)) )  )  'BB+CV+PHU
				if rs("country")="VN" then 					
					if TOTY mod (26*8)<>0 then 
		  				TTMH = fix(TOTY/26/8)+1 		  				
		  			else
		  				TTMH = fix(TOTY/26/8) 
		  			end if 
		  			tmpRec(i, j, 31) = TTMH 
				else
					tmpRec(i, j, 31) = round(tmpRec(i, j, 20)/30,3)
				end if 	
				tmpRec(i, j, 32)=RS("B_QC")
				tmpRec(i, j, 33)= cdbl(TOTY)+cdbl(tmpRec(i, j, 24))+cdbl(tmpRec(i, j, 25))+cdbl(tmpRec(i, j, 26))+cdbl(tmpRec(i, j, 27))+cdbl(tmpRec(i, j, 32))
				tmpRec(i, j, 34)=rs("b_memo")
				if rs("B_BB")="0" then 
					tmpRec(i, j, 35) = "red"
				else
					tmpRec(i, j, 35) = "black"
				end if 
				tmpRec(i, j, 36) = rs("country")
				tmpRec(i, j, 37) = rs("nz")
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
	Session("empsalary01") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empsalary01")
	
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
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	<%=self%>.bbcode(0).focus()	
	'<%=self%>.BBcode(0).SELECT()
end function   

function chgdata()
	<%=self%>.action="empfile.salary.asp?totalpage=0"
	<%=self%>.submit
end function 
-->
</SCRIPT>  
</head>   
<body   topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()"   bgproperties="fixed"  >
<form name="<%=self%>" method="post" action="YECE0101.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>"> 	
<INPUT TYPE=hidden NAME=MMDAYS VALUE="<%=MMDAYS%>">
<INPUT TYPE=hidden NAME=F_country VALUE="<%=F_country%>">

<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">人事薪資系統( 員工薪資管理 ) </TD>		
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<!--------------------------------------------------------------------> 		
<TABLE  CLASS="FONT9" BORDER=0 cellspacing="0" cellpadding="1" > 	
	<tr height=25>		
		<TD colspan=12>計薪年月：<input name=calcYM value="<%=YYMM%>" class=inputbox size=7 maxlength=6> 
		　　<%if F_country="CN" then%><a href="salary_CN_ver200801.pdf" target="_blank"><font color=blue>**查看薪資結構表**</font>(*.pdf 需安裝 Acrobat)</a><%end if%>
		</TD>		 
		 
	</tr>
	<TR HEIGHT=25 BGCOLOR="LightGrey"   >
 		<TD ROWSPAN=2 width=30 align=center>項<BR>次</TD>
 		<TD align=center>工號</TD> 		
 		<TD COLSPAN=3  >員工姓名(中,英,越)</TD>  		
 		<td align=center><%if F_country="CN" then %><%ELSE%>時薪<%END IF%></td>
 		<td align=center>到職日期</td>
 		<td align=center>離職日期</td>
 		<TD align=center><%if F_country="CN" then %><%ELSE%>保險日期<%END IF%></TD>
 		<td align=center colspan=3>備註</td>
 	</TR>
 	<tr BGCOLOR="LightGrey"  HEIGHT=25 > 	
 		<TD align=center>薪資代碼</TD>
 		<TD align=center>基本薪資</TD>
 		<TD align=center>職專</TD> 			
 		<TD align=center>職專加給</TD>	
 		<TD align=center><%if F_country="CN" then %>其他加給<%ELSE%>獎金(Y)<%END IF%></TD>
 		<td align=center>語言加給</td>
 		<td align=center>技術加給</td>
 		<td align=center><%if F_country="CN" then %>年資加給<%else%>環境加給<%end if%></td>
 		<td align=center><%if F_country="CN" then %>補助醫療<%else%>其他加給<%end if%></td>
 		<td align=center>全勤獎金</td>
 		<td align=center>薪資合計</td>
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
 				<font color="<%=tmpRec(CurrentPage, CurrentRow, 35)%>"><%=tmpRec(CurrentPage, CurrentRow, 1)%></font>
 			</a>
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
 			<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
 			<input type=hidden name="COUNTRY" value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>"> 
 		</TD> 		
 		<TD COLSPAN=3>
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<font color="<%=tmpRec(CurrentPage, CurrentRow, 35)%>"><%=tmpRec(CurrentPage, CurrentRow, 2)%>
 				<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font></font>
 			</a>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT TYPE=HIDDEN  NAME=HHMOENY  CLASS='INPUTBOX8' READONLY  SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=HHMOENY >
	 		<%END IF%>	 		
 		</TD> 
 		
 		<TD  ALIGN=CENTER >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=INDAT CLASS='INPUTBOX8' READONLY  SIZE=10 VALUE="<%=(right(tmpRec(CurrentPage, CurrentRow, 5),8))%>(<%=tmpRec(CurrentPage, CurrentRow, 37)%>)" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=INDAT >
	 		<%END IF%> 		
	 	</TD>	
 		<TD  ALIGN=CENTER >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=OUTDAT CLASS='INPUTBOX8' READONLY  SIZE=10 VALUE="<%=RIGHT(tmpRec(CurrentPage, CurrentRow, 30),8)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=OUTDAT >
	 		<%END IF%> 	
 		</TD>
 		<TD  ALIGN=CENTER >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT TYPE=HIDDEN NAME=BHDAT CLASS='INPUTBOX8' READONLY  SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 28)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BHDAT >
	 		<%END IF%> 	
 		</TD>  		
 		<TD colspan=3>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>		
 				<INPUT NAME=memo size=39 class="inputbox" value="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" onchange="DATACHG(<%=CURRENTROW-1%>)">
 			<%else%>	
 				<INPUT TYPE=HIDDEN NAME=memo >
 			<%end if%>
 		</TD>		 	
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> >
 		<TD ALIGN=RIGHT > 			 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			
		 		<select name=BBCODE  class="txt8" style="width:60" onchange="bbcodechg(<%=currentrow-1%>)">						 			
					<%SQL="SELECT * FROM empsalarybasic WHERE FUNC='AA' and country='"& tmpRec(CurrentPage, CurrentRow, 36) &"'   ORDER BY CODE " 
					response.write sql 
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("CODE")%>" <%IF cdbl(RST("bonus"))=cdbl(trim(tmpRec(CurrentPage, CurrentRow, 20))) THEN %> SELECTED <%END IF%> ><%=RST("CODE")%>-<%=RST("bonus")%></option>				 
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
	 			<INPUT NAME=BB CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">	 			
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
 		 		<INPUT NAME=CV CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>"  STYLE="TEXT-ALIGN:RIGHT" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 職務加給"    >
 		 		<input type=hidden name=CVCODE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" SIZE=3>
 		 	<%else%>
				<input type=hidden name=CV >	
				<input type=hidden name=CVCODE >
			<%end if %>	
 		 </TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=PHU CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助獎金(Y)" >
	 		<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=PHU	>
			<%END IF%>	
 		</TD> 		
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME=NN CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 語言加給" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=NN >
			<%END IF%>		
 		</TD>
 		<TD  ALIGN=RIGHT>
	 		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=KT CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 技術加給" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KT >
	 		<%END IF%>			
 		</TD>
 		<TD  ALIGN=RIGHT>
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 
	 			<INPUT NAME=MT CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 環境加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=MT >
	 		<%END IF%>			
 		</TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=TTKH CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TTKH >
	 		<%END IF%>			
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=QC CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 全勤">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QC >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=totamt CLASS='INPUTBOX8' SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 33)%>" STYLE="TEXT-ALIGN:RIGHT; color:darkred;BACKGROUND-COLOR:LIGHTYELLOW"   >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=totamt >
	 		<%END IF%>	
 		</TD> 
 		
	</TR>
	<%next%>
</TABLE>	
<input type=hidden name="empid">
<input type=hidden name="empautoid"> 
<INPUT TYPE=HIDDEN NAME=INDAT > 
<INPUT TYPE=HIDDEN NAME=INDAT > 
<INPUT TYPE=HIDDEN NAME=OUTDAT >
<INPUT TYPE=HIDDEN NAME=BHDAT > 
<INPUT TYPE=HIDDEN NAME=GTDAT > 
<INPUT TYPE=HIDDEN NAME=HHMOENY > 
<input type=hidden name="BBCODE">
<input type=hidden name="BB">
<input type=hidden name="F1_JOB">
<input type=hidden name="CV">
<input type=hidden name="CVCODE">
<input type=hidden name="PHU">
<input type=hidden name="NN">
<input type=hidden name="KT">
<input type=hidden name="MT">
<input type=hidden name="TTKH">
<INPUT TYPE=HIDDEN NAME=QC > 
<INPUT TYPE=HIDDEN NAME=totamt > 
<INPUT TYPE=HIDDEN NAME=memo > 


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
	tmpRec = Session("empsalary01")
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
		tmpRec(CurrentPage, CurrentRow, 33) = request("totamt")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 34) = request("memo")(CurrentRow)		
		
	next 
	Session("empsalary01") = tmpRec
	
End Sub
%> 

<script language=vbscript>
function BACKMAIN() 	
	open "../main.asp" , "_self"
end function   

function clr()
	open "<%=self%>.fore.asp" , "_self"
end function 

function go()	
	<%=self%>.action="<%=self%>.upd.asp"  
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
	open "<%=self%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" & daystr & "&code=" &	codestr , "Back" 		 
	'DATACHG(INDEX)	  
	 
	'PARENT.BEST.COLS="70%,30%"	 	
END FUNCTION 

FUNCTION JOBCHG(INDEX)	
	COUNTRYstr = <%=self%>.COUNTRY(index).value 	
	codestr=<%=self%>.F1_JOB(index).value 
	daystr=<%=self%>.MMDAYS.value 
	'if COUNTRYstr="VN" then 
		open "<%=self%>.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
			 "&days=" &daystr & "&code=" &	codestr , "Back"
	'end if 		 
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
	CODESTR06 = <%=SELF%>.QC(INDEX).VALUE
	CODESTR07 = <%=SELF%>.BB(INDEX).VALUE
	CODESTR08 = <%=SELF%>.CV(INDEX).VALUE	
	CODESTR09 =  (escape(trim(<%=SELF%>.memo(INDEX).VALUE)))
	'daystr=<%=self%>.MMDAYS.value  
	'ALERT CODESTR09
	'ALERT CODESTR03
	
	open "<%=self%>.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 &_
		 "&CODESTR06="& CODESTR06 &_
		 "&CODESTR07="& CODESTR07 &_	
		 "&CODESTR08="& CODESTR08 &_	
		 "&CODESTR09="& CODESTR09  , "Back"  
		 
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

