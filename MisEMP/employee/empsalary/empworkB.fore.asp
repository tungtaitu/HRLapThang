<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<%
SESSION.CODEPAGE="65001"
SELF = "empworkb"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
Set RST = Server.CreateObject("ADODB.Recordset")   

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record   

 
YYMM = REQUEST("YYMM")
IF YYMM="" THEN 	
	YYMM = year(date())&right("00"&month(date()),2)   	
	'YYMM="200601"
	cDatestr=date()
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天 	
ELSE
	cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01") 
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
END IF 	 

 
if month(date())="01" then  
	calcmonth = year(date()-1)&"12" 
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	
EMPID = TRIM(REQUEST("EMPID"))
empautoid = TRIM(REQUEST("empautoid"))

Ftotalpage = request("Ftotalpage")
Fcurrentpage = request("Fcurrentpage")
FRecordInDB = request("FRecordInDB")  
'RESPONSE.END 

'-------------------------------------------------------------------------------------- 
SQL="SELECT b.sys_value as groupstr, c.sys_value as zunostr, d.sys_value as jobstr , a.* from  "&_
	"( SELECT * FROM  EMPFILE WHERE ISNULL(STATUS,'')<>'D' AND autoid='"& empautoid &"' ) a "&_
	"left join ( select * from basicCode where func='groupid' ) b on b.sys_type = a.groupid "&_
	"left join ( select * from basicCode where func='zuno' ) c on c.sys_type = a.zuno "&_
	"left join ( select * from basicCode where func='lev' ) d on d.sys_type = a.job " 
	'RESPONSE.WRITE SQL 
	'RESPONSE.END 
	RST.OPEN SQL , CONN, 3, 3 
IF NOT RST.EOF THEN 
	empautoid = TRIM(RST("AUTOID"))
	EMPID=TRIM(RST("EMPID"))	'員工編號
	INDAT=TRIM(RST("INDAT"))	'到職日  
	TX=TRIM(RST("TX"))	'到職日 
	WHSNO=TRIM(RST("WHSNO"))	'廠別
	UNITNO=TRIM(RST("UNITNO"))	'處/所
	GROUPID=TRIM(RST("GROUPID"))	'組/部門
	ZUNO=TRIM(RST("ZUNO"))	'單位
	JOB=TRIM(RST("JOB"))	'職等	
	EMPNAM_CN=TRIM(RST("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RST("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RST("COUNTRY"))	'國籍 	 
	GROUPSTR = TRIM(RST("GROUPSTR"))  '組/部門
	ZUNOSTR = TRIM(RST("ZUNOSTR"))  '單位
	JOBSTR = TRIM(RST("JOBSTR"))  '職等
END IF 
SET RST=NOTHING  


gTotalPage = 1
'PageRec = 31    'number of records per page
if yymm="" then 
	PageRec = 31 
else 
	PageRec = days 
end if 	
TableRec = 40    'number of fields per record   

'出缺勤紀錄 --------------------------------------------------------------------------------------
'SQLSTRA="select CONVERT(CHAR(10),a.dat,111) AS DAT , a.status,  b.*  "&_
'	 	"from "&_
'		"( SELECT '"& EMPID &"' AS EMPID , * FROM YDBMCALE ) A   "&_
'		"LEFT JOIN  ( SELECT* FROM EMPWORK ) B ON B.WORKDAT = CONVERT(CHAR(8), DAT, 112 ) "&_
'		"where CONVERT(CHAR(6), DAT, 112)='"& YYMM &"' AND A.EMPID='"& EMPID &"'  "&_
'		"order by a.dat, b.empid " 
SQLSTRA= " SP_CALCWORKTIME '"& EMPID &"', '"& YYMM &"' " 
'response.write 	request("TotalPage")  	
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	'RESPONSE.WRITE SQLSTRA 
	'RESPONSE.END 
	rs.Open SQLSTRA, conn, 3, 3 
	IF NOT RS.EOF THEN 
		rs.PageSize = PageRec  		
		RecordInDB = days 'rs.RecordCount 
		TotalPage = 1 'rs.PageCount  
		gTotalPage = TotalPage
	END IF 	 

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
	 for j = 1 to PageRec
		if not rs.EOF then 			
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("DAT"))
			tmpRec(i, j, 2) = RS("T1")
			tmpRec(i, j, 3) = RS("T2") 			
 			'if RS("TIMDIF")< 0 then 
			'	tmpRec(i, j, 4) = (RS("TIMDIF")/2) + 24 
			'else
				tmpRec(i, j, 4) = cdbl(RS("TIMDIF"))
			'end if 	 
			
			  
			'遲到
 			IF  ( tmpRec(i, j, 2)>="08:04" AND  tmpRec(i, j, 2)<="08:15" ) or ( tmpRec(i, j, 2)>="13:04" AND tmpRec(i, j, 2)<="13:15" ) AND ( tmpRec(i, j, 2)>="16:04"  and tmpRec(i, j, 2)<="16:15" ) THEN 
 				'RESPONSE.WRITE tmpRec(i, j, 1)&" "&tmpRec(i, j, 2) &"<br>"
 				'RESPONSE.WRITE tmpRec(i, j, 30) &"<br>"
 				tmpRec(i, j, 31) =  1 
 			ELSE
 				tmpRec(i, j, 31) = 0 		
 			END IF 		 
  
			'RESPONSE.WRITE NDAT1  &"<br>"
			'RESPONSE.WRITE tmpRec(i, j, 30)   &"<br>"
			'RESPONSE.WRITE NDAT2  &"<br>"
			'RESPONSE.WRITE NEWTOTH  &"<br>"
			
			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN 
				'tmpRec(i, j, 5) = 0  	 
				'response.write RS("STATUS") &  rs("groupid") &"<BR>"	
				if RS("STATUS")="H1" then 	
					if rs("groupid")="A061" then 
						if tmpRec(i, j, 4) > 8 then 
							tmpRec(i, j, 5) = tmpRec(i, j, 4) - 8 
						else
							tmpRec(i, j, 5) = 0 
						end if 
					elseif rs("groupid")="A062" or  rs("groupid")="A063" or  rs("groupid")="A067" then  			
						if tmpRec(i, j, 4) > 9.5 then 
							tmpRec(i, j, 5) = tmpRec(i, j, 4) - 9.5 
						else
							tmpRec(i, j, 5) = 0 
						end if 
					else 	
						tmpRec(i, j, 5) = 0  
					end if 	
				else
					tmpRec(i, j, 5)=0	
				end if				
			ELSE
				tmpRec(i, j, 5) = rs("H1")				
			END IF 	 
			
			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN  
				if RS("STATUS")="H2" then 	
					if rs("groupid")="A061" then 
						tmpRec(i, j, 6) = tmpRec(i, j, 4)  					
					elseif rs("groupid")="A062" or  rs("groupid")="A063" or  rs("groupid")="A067" then  			
						if tmpRec(i, j, 4) > 8 then 
							tmpRec(i, j, 6) = tmpRec(i, j, 4) - 1 
						else
							tmpRec(i, j, 6) = 0 
						end if  
					else 	
						tmpRec(i, j, 6) = 0  
					end if 	
				else
					tmpRec(i, j, 6)=0		
				end if 
			ELSE
				tmpRec(i, j, 6) = rs("H2")
			END IF 	
			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN  
				if RS("STATUS")="H3" then 	
					if rs("groupid")="A061" then 
						tmpRec(i, j, 7) = tmpRec(i, j, 4)  					
					elseif rs("groupid")="A062" or  rs("groupid")="A063" or  rs("groupid")="A067" then  			
						if tmpRec(i, j, 4) > 8 then 
							tmpRec(i, j, 7) = tmpRec(i, j, 4) - 1 
						else
							tmpRec(i, j, 7) = 0 
						end if  
					else
						tmpRec(i, j, 7) = 0  	
					end if 	
				else
					tmpRec(i, j, 7)=0		
				end if 
			ELSE
				tmpRec(i, j, 7) = rs("H3")
			END IF 	
			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN 				
				IF LEFT(RS("T2"),2) >="22" OR ( LEFT(RS("T2"),2) >="00"  AND LEFT(RS("T2"),2) <="05"  )  THEN
					IF  LEFT(RS("T1"),2) >="22" OR ( LEFT(RS("T1"),2) >="00"  AND LEFT(RS("T1"),2) <="05"  ) THEN 
						NNT1 = tmpRec(i, j, 1)&" "&RS("NEWT1")
					ELSE 	
						NNT1 = tmpRec(i, j, 1)&" 22:00"	
					END IF 
					NNT2 = tmpRec(i, j, 1)&" "&RS("NEWT2")
					'RESPONSE.WRITE NNT1 &"<br>"
					'RESPONSE.WRITE NNT2 &"<br>"
					IF NNT2<NNT1 THEN 
						tmpRec(i, j, 8) = ROUND( DATEDIFF("N", NNT1, NNT2)/30 ,0) / 2  + 24
					ELSE
						tmpRec(i, j, 8) = ROUND( DATEDIFF("N", NNT1, NNT2)/30 ,0) / 2  
					END IF 	
				ELSE
					tmpRec(i, j, 8) = 0 		
				END IF 
			ELSE
				tmpRec(i, j, 8) = rs("B3")
			END IF 	
			IF ISNULL(RS("JIAA")) THEN 
				tmpRec(i, j, 9) = 0 
			ELSE
				tmpRec(i, j, 9) = rs("JIAA")
			END IF 	
			IF ISNULL(RS("JIAB")) THEN 
				tmpRec(i, j, 10) = 0 
			ELSE
				tmpRec(i, j, 10) = rs("JIAB")
			END IF 	
			IF ISNULL(RS("JIAC")) THEN 
				tmpRec(i, j, 11) = 0 
			ELSE
				tmpRec(i, j, 11) = rs("JIAC")
			END IF 	
			IF ISNULL(RS("JIAD")) THEN 
				tmpRec(i, j, 12) = 0 
			ELSE
				tmpRec(i, j, 12) = rs("JIAD")
			END IF 	
			IF ISNULL(RS("JIAE")) THEN 
				tmpRec(i, j, 13) = 0 
			ELSE
				tmpRec(i, j, 13) = rs("JIAE")
			END IF 	
			IF ISNULL(RS("JIAF")) THEN 
				tmpRec(i, j, 14) = 0 
			ELSE
				tmpRec(i, j, 14) = rs("JIAF")
			END IF 	 		
			IF ISNULL(RS("JIAG")) THEN 
				tmpRec(i, j, 21) = 0 
			ELSE
				tmpRec(i, j, 21) = cdbl(rs("JIAG"))
			END IF 	  
			
			tmpRec(i, j, 15)= mid("日一二三四五六",weekday(tmpRec(i, j, 1)) , 1 )	 	
			
			if rs("T1")<>"" then 
				tmpRec(i, j, 16) = "readonly"
			else
				tmpRec(i, j, 16) = "inputbox"
			end if 	
			if rs("T2")<>"" then 
				tmpRec(i, j, 17) = "readonly"
			else
				tmpRec(i, j, 17) = "inputbox"
			end if 		 
			
			tmpRec(i, j, 18)=RS("STATUS") 	
			'所有假加總		
			tmpRec(i, j, 22) =0 'cdbl(tmpRec(i, j, 9)+cdbl(tmpRec(i, j, 10))+cdbl(tmpRec(i, j, 11))+cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 21)))
			'曠職
			if  cdate(trim(rs("DAT"))) <= cdate(date()) and cdate(trim(rs("DAT")))>= cdate(trim(INDAT)) and tmpRec(i, j, 18)="H1" and  ( cdbl(tmpRec(i, j, 4))=0  or  ( cdbl(tmpRec(i, j, 4))+ tmpRec(i, j, 22) )< 8  )   then 
				tmpRec(i, j, 19) = 8 - cdbl(tmpRec(i, j, 4)) - tmpRec(i, j, 22) 
				'response.write "A" &"<BR>"
			else
				tmpRec(i, j, 19) = 0 
				'response.write "B" &"<BR>"
			end if
			'忘刷
			if  cdate(trim(rs("DAT"))) <= cdate(date()) and tmpRec(i, j, 18)="H1"  and (tmpRec(i, j, 2)<>"" and tmpRec(i, j, 3)<>"" )  and ( cdbl(tmpRec(i, j, 4))=0  and  ( cdbl(tmpRec(i, j, 4))+tmpRec(i, j, 22) ) < 8 )   then 
				tmpRec(i, j, 20) = tmpRec(i, j, 20) + 1
				'response.write "A" &"<BR>"
			else
				tmpRec(i, j, 20) = tmpRec(i, j, 20) + 0 
				'response.write "B" &"<BR>"
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
	Session("empworkb") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empworkb")
	
	Select case request("send") 
	     Case "FIRST"
		      CurrentPage = 1			
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then 
			     CurrentPage = CurrentPage - 1				
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) <= cint(TotalPage) then 
			     CurrentPage = CurrentPage + 1 
		      end if			
	     Case "END"
		      CurrentPage = TotalPage 			
	     Case Else 
		      CurrentPage = 1	
	end Select 
end if    

 
  
'--------------------------------------------------------------------------------------  
FUNCTION FDT(D)
IF D <> "" THEN
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2) 	
END IF 	
END FUNCTION 
'--------------------------------------------------------------------------------------
SQL="SELECT * FROM BASICCODE WHERE FUNC='CLOSEP' AND SYS_TYPE='"& YYMM &"' " 
SET RDS=CONN.EXECUTE(SQL) 
IF RDS.EOF THEN 
	PCNTFG = 1 '可異動
	MSGSTR=""
ELSE
	PCNTFG = 0 '不可異動該月出勤紀錄	
	MSGSTR="已結算，不可異動"
END IF 
SET RDS=NOTHING 
IF PCNTFG = "0" THEN 
	INPUTSTS="READONLY" 
ELSE
	INPUTSTS="INPUTBOX"	
END IF 
'---------------------------------------------------------------------------------
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 			
end function  

function f()
	<%=self%>.TIMEUP(0).SELECT()		
end function   

function colschg(index) 
	thiscols = document.activeElement.name 	 	   
	if window.event.keyCode = 38 then  	
		IF INDEX<>0 THEN 
			document.all(thiscols)(index-1).SELECT()
		END IF 	
	end if  
	if window.event.keyCode = 40 then  
		document.all(thiscols)(index+1).SELECT() 		
	end if  

end function  

-->
</SCRIPT>  
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"    >
<form name="<%=self%>"  method="post"  >
<INPUT TYPE=HIDDEN NAME="PCNTFG" VALUE=<%=PCNTFG%>>	
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>	
<INPUT TYPE=HIDDEN NAME="empautoid" VALUE=<%=empautoid%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>"> 
<INPUT TYPE=hidden NAME=FTotalPage VALUE="<%=FTotalPage%>"> 
<INPUT TYPE=hidden NAME=FCurrentPage VALUE="<%=FCurrentPage%>"> 
<INPUT TYPE=hidden NAME=FRecordInDB VALUE="<%=FRecordInDB%>"> 
<!-- table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<TD align=center >員工差勤作業 </TD>		
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500 -->		
<table width=500   class=font9>
	<TR>
		<td >查詢年月:</td>
		<td COLSPAN=3>
			<select name=yymm class=font9  onchange="dchg()" disabled >
				<%for z = 1 to 12 
				  yymmvalue = year(date())&right("00"&z,2)
				%>
					<option value="<%=yymmvalue%>" <%if yymmvalue=yymm then %>selected<%end if%>><%=yymmvalue%></option>
				<%next%>	
			</select>  				
			<input type=hiddenT class=readonly readonly  name=days value="<%=days%>" size=5>
			　<FONT COLOR=RED><%=MSGSTR%></FONT>
		</td> 
	</TR>
	<tr height=30>					
		<td width=60>員工編號:</td>
		<td>
			<input name=empid value="<%=EMPID%>" size=7 class="readonly" readonly style="height:22">
			<input name=empnam value="<%=empnam_cn&" "&empnam_vn%>" size=30 class="readonly" readonly style="height:22">
		</td>
		<td>單位:</td>
		<td>
			<input name=groupidstr value="<%=GROUPSTR%>" size=7 class="readonly" readonly  style="height:22">
			<input name=zunostr value="<%=zunoSTR%>" size=5 class="readonly" readonly style="height:22" >
			<input TYPE=HIDDEN name=groupid value="<%=groupid%>" size=5 >
		</td>
	</tr>		
</table>  
<table width=500 class=font9 >
	<tr>		
		<td width=60>到職日期:</td>
		<td><input name=indat value="<%=indat%>" size=11 class="readonly" readonly  style="height:22"></td>
		
		<td>職等:</td>
		<td><input name=job value="<%=jobSTR%>" size=12 class="readonly" readonly  style="height:22"></td>
		<td>特休(天/小時):</td>
		<td>
			<input name=TX value="<%=tx%>" size=5 class="readonly" readonly  style="height:22">
			<input name=TXH value="<%=tx*8%>" size=5 class="readonly" readonly  style="height:22">
		</td>
	</tr>
</table>	 
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>	
<TABLE WIDTH=610 CLASS=FONT9 >
	<TR BGCOLOR=#CCCCCC>
		<TD ROWSPAN=2 ALIGN=CENTER>日期</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>上班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>下班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>工時</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>曠職</TD>	
		<TD ROWSPAN=2 ALIGN=CENTER>忘<br>刷<br>卡</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>遲到</TD>	
		<TD COLSPAN=4 ALIGN=CENTER>加班(單位：小時)</TD>
		<TD COLSPAN=7 ALIGN=CENTER>休假(單位：小時)</TD>
	</TR>
	<TR BGCOLOR=#CCCCCC>
		<TD ALIGN=CENTER>一般(1.5)</TD>
		<TD ALIGN=CENTER>休息(2)</TD>
		<TD ALIGN=CENTER>假日(3)</TD>
		<TD ALIGN=CENTER>夜班(0.3)</TD>
		<TD ALIGN=CENTER>公假</TD>
		<TD ALIGN=CENTER>年假</TD>
		<TD ALIGN=CENTER>事假</TD>
		<TD ALIGN=CENTER>病假</TD>
		<TD ALIGN=CENTER>婚假</TD>
		<TD ALIGN=CENTER>喪假</TD>		
		<TD ALIGN=CENTER>產假</TD>
	</TR>	
	<%
	sum_TOTHOUR = 0
	sum_KZhour = 0  
	sum_Forget = 0 
	sum_H1 = 0
	sum_H2 = 0
	sum_H3 = 0
	sum_B3 = 0
	um_JIAA = 0 
	sum_JIAB = 0		
	sum_JIAC = 0
	sum_JIAD = 0
	sum_JIAE = 0
	sum_JIAF = 0 
	sum_JIAG = 0
	sum_LATEFOR = 0
	
	for CurrentRow = 1 to PageRec	 
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR=""
		END IF 	  
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
	%> 
	<TR>
		<TD ALIGN=CENTER NOWRAP >
		<INPUT NAME=WORKDATIM VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow, 15)&")"%>" CLASS=READONLY READONLY  SIZE=15 STYLE="TEXT-ALIGN:CENTER">
		<INPUT TYPE=HIDDEN NAME=WORKDAT VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" >
		<INPUT TYPE=HIDDEN NAME=STATUS VALUE="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" >
		</TD>
		<TD ALIGN=CENTER><INPUT NAME=TIMEUP VALUE="<%=tmpRec(CurrentPage, CurrentRow, 2)%>" CLASS=<%=tmpRec(CurrentPage, CurrentRow, 16)%> SIZE=6 STYLE="TEXT-ALIGN:CENTER"  maxlength=4  readonly ></TD>
		<TD ALIGN=CENTER><INPUT NAME=TIMEDOWN VALUE="<%=tmpRec(CurrentPage, CurrentRow, 3)%>" CLASS=<%=tmpRec(CurrentPage, CurrentRow, 17)%> SIZE=6 STYLE="TEXT-ALIGN:CENTER" maxlength=4  readonly  ></TD>
		<TD ALIGN=CENTER><INPUT NAME=TOTHOUR VALUE="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"  onkeydown="colschg(<%=CurrentRow-1%>)" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=KZhour VALUE="<%=tmpRec(CurrentPage, CurrentRow, 19)%>" CLASS="readonly" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 19)<>"0" then %>red<%else%>black<%end if%>" readonly   ></TD>		
		<TD ALIGN=CENTER><INPUT NAME=Forget VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" CLASS=<%=INPUTSTS%> SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 20)<>"0" then %>red<%else%>black<%end if%>" readonly ></TD>
		<TD ALIGN=CENTER><INPUT NAME=LATEFOR VALUE="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" CLASS=<%=INPUTSTS%> SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 31)<>"0" then %>red<%else%>black<%end if%>" readonly  ></TD>
		<TD ALIGN=CENTER><INPUT NAME=H1 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 5)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 5)<>"0" then %>red<%else%>black<%end if%>" readonly ></TD>
		<TD ALIGN=CENTER><INPUT NAME=H2 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 6)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 6)<>"0" then %>red<%else%>black<%end if%>" readonly ></TD>
		<TD ALIGN=CENTER><INPUT NAME=H3 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 7)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 7)<>"0" then %>red<%else%>black<%end if%>" readonly ></TD>
		<TD ALIGN=CENTER><INPUT NAME=B3 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 8)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 8)<>"0" then %>red<%else%>black<%end if%>" readonly  ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAG VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 21)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 13)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 13)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAA VALUE="<%=tmpRec(CurrentPage, CurrentRow, 9)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 9)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAB VALUE="<%=tmpRec(CurrentPage, CurrentRow, 10)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 10)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAC VALUE="<%=tmpRec(CurrentPage, CurrentRow, 11)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 11)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAD VALUE="<%=tmpRec(CurrentPage, CurrentRow, 12)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 12)<>"0" then %>red<%else%>black<%end if%>" ></TD>		
		<TD ALIGN=CENTER><INPUT NAME=JIAF VALUE="<%=tmpRec(CurrentPage, CurrentRow, 14)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 14)<>"0" then %>red<%else%>black<%end if%>"  ></TD>		
	</TR>	
	<%
		sum_TOTHOUR = sum_TOTHOUR + cdbl(tmpRec(CurrentPage, CurrentRow, 4))
		sum_LATEFOR  = sum_LATEFOR + cdbl(tmpRec(CurrentPage, CurrentRow, 31)) 		
		sum_KZhour  = sum_KZhour + cdbl(tmpRec(CurrentPage, CurrentRow, 19))
		sum_Forget  = sum_Forget + cdbl(tmpRec(CurrentPage, CurrentRow, 20))
		sum_H1 = sum_H1 + cdbl(tmpRec(CurrentPage, CurrentRow, 5))
		sum_H2 = sum_H2 + cdbl(tmpRec(CurrentPage, CurrentRow, 6))
		sum_H3 = sum_H3 + cdbl(tmpRec(CurrentPage, CurrentRow, 7))
		sum_B3 = sum_B3 + cdbl(tmpRec(CurrentPage, CurrentRow, 8)) 
		sum_JIAA = sum_JIAA + cdbl(tmpRec(CurrentPage, CurrentRow, 9)) 
		sum_JIAB = sum_JIAB	+ cdbl(tmpRec(CurrentPage, CurrentRow, 10)) 	
		sum_JIAC = sum_JIAC + cdbl(tmpRec(CurrentPage, CurrentRow, 11))
		sum_JIAD = sum_JIAD + cdbl(tmpRec(CurrentPage, CurrentRow, 12))
		sum_JIAE = sum_JIAE + cdbl(tmpRec(CurrentPage, CurrentRow, 13))
		sum_JIAF = sum_JIAF + cdbl(tmpRec(CurrentPage, CurrentRow, 14))
		sum_JIAG = sum_JIAG + cdbl(tmpRec(CurrentPage, CurrentRow, 21))
	%>
	<%next%> 	 
	<tr BGCOLOR="Lavender" >
		<td align=right colspan=3 HEIGHT=22>總計</td>
		<td align=right ><INPUT NAME="sum_TOTHOUR" VALUE="<%=sum_TOTHOUR%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_KZhour" VALUE="<%=sum_KZhour%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>		
		<td align=right ><INPUT NAME="sum_Forget" VALUE="<%=sum_Forget%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_LATEFOR" VALUE="<%=sum_LATEFOR%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>		
		<td align=right ><INPUT NAME="sum_H1" VALUE="<%=sum_H1%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_H2" VALUE="<%=sum_H2%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_H3" VALUE="<%=sum_H3%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_B3" VALUE="<%=sum_B3%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_JIAG" VALUE="<%=sum_JIAG%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>		
		<td align=right ><INPUT NAME="sum_JIAE" VALUE="<%=sum_JIAE%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_JIAA" VALUE="<%=sum_JIAA%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_JIAB" VALUE="<%=sum_JIAB%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_JIAC" VALUE="<%=sum_JIAC%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>
		<td align=right ><INPUT NAME="sum_JIAD" VALUE="<%=sum_JIAD%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>		
		<td align=right ><INPUT NAME="sum_JIAF" VALUE="<%=sum_JIAF%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"></td>		
	</tr>
</TABLE>
	
<TABLE border=0 width=600 class=font9 >
<tr>
    <td align="CENTER" height=40  >     
	<input type=BUTTON name=send value="關閉此視窗(CLOSE)"  class=button ONCLICK="vbscript:window.close()">　　	 
	</td>
	 
</TR>
</TABLE> 
  
</form>


</body>
</html>

<script language=vbscript >
 
 

'*******檢查日期*********************************************
FUNCTION date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.BHDAT.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.pduedate.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.vduedate.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.outdat.value)			
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat.value=ANS
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.vduedate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=""
			Document.<%=self%>.BHDAT.focus()
		elseif a=3 then
			Document.<%=self%>.pduedate.value=""
			Document.<%=self%>.pduedate.focus()
		elseif a=4 then
			Document.<%=self%>.vduedate.value=""
			Document.<%=self%>.vduedate.focus()
		elseif a=5 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
		end if		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION

'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("00" & Month(d), 2) & "/" & _
		Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________ 

FUNCTION CHKVALUE(N) 
IF N=1 THEN 
	IF TRIM(<%=SELF%>.BYY.VALUE)<>"" THEN 
		IF ISNUMERIC(<%=SELF%>.BYY.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BYY.VALUE,"-")>0 THEN 
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BYY.VALUE=""
			<%=SELF%>.BYY.FOCUS()
			EXIT FUNCTION 
		END IF 
	END IF 	
ELSEIF N=2 THEN 
	IF TRIM(<%=SELF%>.BMM.VALUE)<>"" THEN 
		IF ISNUMERIC(<%=SELF%>.BMM.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BMM.VALUE,"-")>0 THEN 
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BMM.VALUE=""
			<%=SELF%>.BMM.FOCUS()
			EXIT FUNCTION 
		END IF 
	END IF 
ELSEIF N=3 THEN 
	IF TRIM(<%=SELF%>.BDD.VALUE)<>"" THEN 
		IF ISNUMERIC(<%=SELF%>.BDD.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BDD.VALUE,"-")>0 THEN 
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BDD.VALUE=""
			<%=SELF%>.BDD.FOCUS()
			EXIT FUNCTION 
		END IF 
	END IF 
ELSEIF N=4 THEN 
	IF TRIM(<%=SELF%>.AGES.VALUE)<>"" THEN 
		IF ISNUMERIC(<%=SELF%>.AGES.VALUE)=FALSE OR INSTR(1,<%=SELF%>.AGES.VALUE,"-")>0 THEN 
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.AGES.VALUE=""
			<%=SELF%>.AGES.FOCUS()
			EXIT FUNCTION 
		END IF 
	END IF 
ELSEIF N=5 THEN 
	IF TRIM(<%=SELF%>.GTDAT.VALUE)<>"" THEN 
		IF ISNUMERIC(<%=SELF%>.GTDAT.VALUE)=FALSE OR INSTR(1,<%=SELF%>.GTDAT.VALUE,"-")>0 THEN 
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.GTDAT.VALUE=""
			<%=SELF%>.GTDAT.FOCUS()
			EXIT FUNCTION 
		END IF 
	END IF  		
END IF 

END FUNCTION 

</script>


