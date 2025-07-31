<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
SESSION.CODEPAGE="65001"
SELF = "YEDE02"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set RST = Server.CreateObject("ADODB.Recordset")



D1=request("D1")
D2=request("D2")
F_whsno=request("F_whsno")
F_empid = request("F_empid")
F_Groupid = request("F_groupid")
F_shift = request("F_shift")
F_country = request("F_country")

dd = datediff("d",cdate(D1), cdate(D2))+1 

gTotalPage = 1
PageRec = dd*10    'number of records per page
TableRec = 35    'number of fields per record 

'response.write yymm
IF YYMM="" THEN
	YYMM = year(date())&right("00"&month(date()),2)
	'YYMM="200601"
	cDatestr=date()
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
ELSE
	cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
END IF

nowmonth = year(date())&right("00"&month(date()),2) 
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
EMPID = TRIM(REQUEST("EMPID"))
 
'RESPONSE.END 
 	 
 

'出缺勤紀錄 --------------------------------------------------------------------------------------

sql="select  FG.lsempid, B.EMPID AS Nempid,  c.status as dsts, convert(char(10),c.dat,111) as dat,  b.empnam_cn,  b.empnam_vn , b.shift, a.* from  "&_ 
	"(select 't' as tmp,  * from  ydbmcale where convert(char(8), dat, 112) between '"& replace(D1,  "/","") &"' and '"& replace(D2,  "/","")&"' ) c  "&_	
	"left join (select 't' as tmp, * from  view_empfile where empid like '"& f_empid &"%' )  b on b.tmp = c.tmp "&_	
	"left join (select  *  , isB3 = case when left(timeup,2)>='17' and left(timedown,1)>='0' then 'N'  else  '' end  from  empwork where  workdat between '"& replace(D1,  "/","") &"' and '"& replace(D2,  "/","")&"' )  a on convert(char(8),c.dat,112)=a.workdat and a.empid = b.empid  "&_  	
	"left join (select  empid,ltrim(rtrim(isnull(lsempid,''))) lsempid, convert(char(8), dat, 112) fgdat , sum(toth) as fgH,  sum( case when ltrim(rtrim(lsempid)) ='' then 1 else 0 end  ) fgcnt  from  empforget  "&_
	"where isnull(status,'')<>'D'  group by  empid, convert(char(8), dat, 112),ltrim(rtrim(isnull(lsempid,''))) ) FG on fg.empid = b.empid and fg.fgdat = convert(char(8), c.dat, 112)   "&_
	"where  convert(char(8), dat, 112) >= convert(char(8), b.indat, 112)  "&_
	"and ( isnull(b.outdat,'')='' or isnull(b.outdat,'')<>'' and a.workdat < convert(char(8), b.outdat, 112) ) "&_
	"and b.country like '"& F_country &"%' and  b.groupid like '"& f_groupid  &"%' "&_ 
	"and b.whsno like  '"& f_whsno  &"%'   "&_ 
	"and b.empid like '"& f_empid &"%'  " 
if  f_shift="M"  then 
	sql=sql& "  and ( isnull(a.b3,0)>0 and isnull(a.b3,0)<>7 )  and isB3<>''  "
else 
	sql=sql& " and charindex(case when '"&f_shift&"' ='' then ',' else  '"&f_shift&"' end ,','+isB3+isnull(b.shift,'') ) >0 "
end if  

sql=sql&" order by b.groupid, b.shift, a.empid ,  a.workdat " 	

'response.write sql 
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	'RESPONSE.WRITE SQLSTRA
	'RESPONSE.END
	rs.Open SQL, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF
	'response.write days 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1)	= trim(rs("dat"))						
			 
			if isnull(rs("timeup")) then 
				tmpRec(i, j, 2)=""
			elseif RS("timeup")="000000" then 
				tmpRec(i, j, 2)=""
			else
				tmpRec(i, j, 2)= left(RS("timeup"),2)&":"&mid(RS("timeup"),3,2)
			end if 	
				
			if isnull(rs("timedown"))  then 
				tmpRec(i, j, 3) = ""
			elseif rs("timedown")="000000" then 
				tmpRec(i, j, 3) = ""
			else
				tmpRec(i, j, 3) = left(RS("timedown"),2)&":"&mid(RS("timedown"),3,2)  			
			end if 	
			tmpRec(i, j, 4) = rs("nempid")
			tmpRec(i, j, 5) = rs("empnam_cn")
			tmpRec(i, j, 6) = rs("empnam_vn")			
			tmpRec(i, j, 7) = rs("toth")						
			tmpRec(i, j, 8)= mid("日一二三四五六",weekday(cdate(rs("dat"))) , 1 ) 
			tmpRec(i, j, 9)  = trim(rs("forget"))
			if ( isnull(rs("workdat")) or (RS("timeup")="000000" and rs("timedown")="000000") ) and rs("dsts")="H1"   then 
				tmpRec(i, j, 10)  = 8 
			else
				tmpRec(i, j, 10)  = rs("kzhour")				
			end if 
				
			tmpRec(i, j, 11)="0"			
			tmpRec(i, j, 12) = rs("lateFor")
			tmpRec(i, j, 13) = rs("h1")
			tmpRec(i, j, 14) = rs("h2")
			tmpRec(i, j, 15) = rs("h3")
			tmpRec(i, j, 16) = rs("b3")			
			tmpRec(i, j, 17) = rs("jiaG")
			tmpRec(i, j, 18) = rs("jiaE")
			tmpRec(i, j, 19) = rs("jiaA")
			tmpRec(i, j, 20) = rs("jiaB")
			tmpRec(i, j, 21) = rs("jiaC")
			tmpRec(i, j, 22) = rs("jiaD")
			tmpRec(i, j, 23) = rs("jiaF")
			tmpRec(i, j, 24) = rs("jiaH")
			if rs("dsts")="H1" then 
				tmpRec(i, j, 25) = rs("H1")
			elseif 	rs("dsts")="H2" then  
				tmpRec(i, j, 25) = rs("H2")
			elseif rs("dsts")="H3" then 
				tmpRec(i, j, 25) = rs("H3")
			else
				tmpRec(i, j, 25) = 0 
			end if 
			tmpRec(i, j, 26) = rs("dsts") 
			tmpRec(i, j, 27) = rs("BC") 
			tmpRec(i, j, 28) = rs("shift") 
			if isnull(rs("jiaa")) then jiaa= 0 else jiaa=rs("jiaa")
			if isnull(rs("jiaB")) then jiaB= 0 else jiaB=rs("jiaB")
			if isnull(rs("jiaC")) then jiaC= 0 else jiaC=rs("jiaC")
			if isnull(rs("jiaD")) then jiaD= 0 else jiaD=rs("jiaD")
			if isnull(rs("jiaE")) then jiaE= 0 else jiaE=rs("jiaE")
			if isnull(rs("jiaF")) then jiaF= 0 else jiaF=rs("jiaF")
			if isnull(rs("jiaG")) then jiaG= 0 else jiaG=rs("jiaG")
			if isnull(rs("jiaH")) then jiaH= 0 else jiaH=rs("jiaH")
			
			tmpRec(i, j, 29) = cdbl(jiaA)+cdbl(jiaB)+cdbl(jiaC)+cdbl(jiaD)+cdbl(jiaE)+cdbl(jiaF)+cdbl(jiaG)+cdbl(jiaH) 
			
			if rs("dsts")="H1" then 
				tmpRec(i, j, 30)="Black"
				tmpRec(i, j, 31)="LavenderBlush"
			else	
				tmpRec(i, j, 30)="Crimson" 
				tmpRec(i, j, 31)="#e4e4e4"
			end if 	
			'response.write tmpRec(i, j, 9) &"<BR>"
			tmpRec(i, j, 32) = rs("lsempid")  '臨時卡
			
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
	Session("YEDE02B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))	
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	StoreToSession()	
	tmpRec = Session("YEDE02B")

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
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function 

function gos()
	<%=self%>.totalpage.value="0"
	<%=self%>.submit()
end function 
 
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0" onkeydown="enterto()"  >
<form name="<%=self%>"  method="post" action = "<%=self%>.ForeGnd.asp" >
<INPUT TYPE=HIDDEN NAME="PCNTFG" VALUE="<%=PCNTFG%>">
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>"> 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3>
					<Tr>
						<td width=50 align=right>日期:</td>
						<td nowrap>
							<input type="text" style="width:100px" name=D1 maxlength=10 onblur="date_change(1)" value="<%=DD2%>">~
							<input type="text" style="width:100px" name=D2 maxlength=10 onblur="date_change(2)" value="<%=DD2%>">
						</td> 
						<TD  align=right nowrap >工號:</TD>
						<td colspan=7>
							<input type="text" style="width:100px" name=F_empid maxlength=5  value="<%=F_empid%>">
						</td>
					</tr>
					<tr>	
						<TD  align=right nowrap>國籍:</TD>
						<TD >
							<select name=F_country     onchange="gos()">							
							<option value="">全部 </option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc "			  
							  SET RST = CONN.EXECUTE(SQL)
							  WHILE NOT RST.EOF	%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_country then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>
							  <%RST.MOVENEXT
								WEND
								SET RST=NOTHING %>
							</SELECT>				
						</TD>		
						<td align=right nowrap>廠別:</td>
						<TD >
							<select name=F_whsno    onchange="gos()" style='width:120px' >							
							<option value="">全部 </option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='whsno'  ORDER BY SYS_type desc "			  
							  SET RST = CONN.EXECUTE(SQL)
							  WHILE NOT RST.EOF	%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
							  <%RST.MOVENEXT
								WEND
								SET RST=NOTHING %>
							</SELECT>				
						</TD>			
						<td align=right nowrap>部門:</td>
						<TD >
							<select name=F_groupid    onchange="gos()" style='width:120px' >							
							<option value="">全部 </option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='groupid'  ORDER BY SYS_type desc "			  
							  SET RST = CONN.EXECUTE(SQL)
							  WHILE NOT RST.EOF	%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_groupid then%>selected<%end if%>><%=RST("SYS_type")%>-<%=RST("SYS_VALUE")%></option>
							  <%RST.MOVENEXT
								WEND
								SET RST=NOTHING %>
							</SELECT>				
						</TD>			
						<td align=right nowrap>班別:</td>
						<TD>
							<select name=F_shift   onchange="gos()" style="width:120px" >							
							<option value="">全部 </option>
							<option value="N" <%if F_shift="N" then%>selected<%end if%>>ca diem夜班</option>					
							<option value="M" <%if F_shift="M" then%>selected<%end if%>>ca dem < 7H 夜班 </option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='shift'  ORDER BY sys_value desc "			  
							  SET RST = CONN.EXECUTE(SQL)
							  WHILE NOT RST.EOF	
							%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_shift then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>
							 <%
							RST.MOVENEXT
							WEND
							SET RST=NOTHING 
							%>
							</SELECT>				
						</TD>
						<td>
							<input type=button name="btn" value="(S)查詢" class="btn btn-sm btn-outline-secondary" onclick="gos()">
						</td>			
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" style="width:98%">
					<TR BGCOLOR=#e4e4e4 >
						<TD ROWSPAN=2 ALIGN=CENTER width=85 nowrap  >日期</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>工號</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=65 nowrap>姓名</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=30 nowrap>班別</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=30 nowrap>排班</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>上班</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>下班</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>工時</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>曠職</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=30 nowrap>忘刷卡</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=30 nowrap>遲到</TD>
						<TD COLSPAN=2 ALIGN=CENTER   nowrap>加班(單位：H)</TD>
						<TD ROWSPAN=2 ALIGN=CENTER width=50 nowrap>臨時卡</TD>
						<TD COLSPAN=8 ALIGN=CENTER   nowrap>休假(單位：小時)</TD>
					</TR>
					<TR BGCOLOR=#e4e4e4 >
						<TD ALIGN=CENTER width=30 nowrap>加班</TD>		
						<TD ALIGN=CENTER width=30 nowrap>夜班(0.3)</TD>
						<TD ALIGN=CENTER width=30 nowrap>公假(G)</TD>
						<TD ALIGN=CENTER width=30 nowrap>年假(E)</TD>
						<TD ALIGN=CENTER width=30 nowrap>事假(A)</TD>
						<TD ALIGN=CENTER width=30 nowrap>病假(B)</TD>
						<TD ALIGN=CENTER width=30 nowrap>婚假(C)</TD>
						<TD ALIGN=CENTER width=30 nowrap>喪假(D)</TD>
						<TD ALIGN=CENTER width=30 nowrap>產假(F)</TD>
						<TD ALIGN=CENTER width=30 nowrap>工傷(H)</TD>
					</TR>
					<%
					sum_TOTHOUR = 0
					sum_KZhour = 0
					sum_Forget = 0
					sum_H1 = 0
					sum_H2 = 0
					sum_H3 = 0
					sum_B3 = 0
					sum_JIAA = 0
					sum_JIAB = 0
					sum_JIAC = 0
					sum_JIAD = 0
					sum_JIAE = 0
					sum_JIAF = 0
					sum_JIAG = 0
					sum_JIAH = 0
					sum_LATEFOR = 0

					for CurrentRow = 1 to PageRec
					'response.write  PageRec &"<BR>"		
					'response.end 
					
					IF CurrentRow MOD 2 = 0   THEN
						WKCOLOR="PaleGoldenrod"  '"LavenderBlush"
					ELSE
						WKCOLOR="PaleGoldenrod"		
					END IF	 
					
					if tmpRec(CurrentPage, CurrentRow, 26)<>"H1" then 
						WKCOLOR="#e4e4e4"		
					end if 
					if tmpRec(CurrentPage, CurrentRow, 4)<>"" then 
					%>
					<%if CurrentRow>1 and tmpRec(CurrentPage, CurrentRow-1, 5)<> tmpRec(CurrentPage, CurrentRow, 5) then%>
						<tr>
							<td colspan=22 height=2><hr size=0	style='border: 1px dotted #999999;' align=left width="100%"></td>
						</tr>
					<%end if%>	
					<TR BGCOLOR=<%=WKCOLOR%> >
						<TD ALIGN=CENTER NOWRAP >
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>
								<font color="<%=tmpRec(CurrentPage, CurrentRow, 30)%>">
								<%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow,8)&")"%>
								</font>
							<%end if%>
							<input type=hidden name="op" value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>" 8 size=1>
							<input type=hidden name="dsts" value="<%=tmpRec(CurrentPage, CurrentRow, 26)%>">
							<input type=hidden name="BC" value="<%=tmpRec(CurrentPage, CurrentRow, 27)%>">
							
						</TD>
						<TD ALIGN=CENTER>
							<a href="vbscript:showWorkTime(<%=currentrow-1%>)"><font color=blue><u><%=tmpRec(CurrentPage, CurrentRow, 4)%></u></font></a><br>			
							<input type=hidden name="empid" value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>">
							<input type=hidden name="workdat" value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
							<input type=hidden name="TotJia" value="<%=tmpRec(CurrentPage, CurrentRow, 29)%>">
						</TD>
						<TD ALIGN=left>			
								<a href="vbscript:showTime(<%=currentrow-1%>)">
									<%=tmpRec(CurrentPage, CurrentRow, 5)%><br>
									<%=left(tmpRec(CurrentPage, CurrentRow, 6),13)%>
								</a>			
						</TD>
						<TD ALIGN=CENTER>
							<%=tmpRec(CurrentPage, CurrentRow, 28)%>
						</TD>
						<TD ALIGN=CENTER>
							<%=tmpRec(CurrentPage, CurrentRow, 27)%>
						</TD>
						<TD ALIGN=CENTER> 
							<input type="text" name="T1"  value="<%=tmpRec(CurrentPage, CurrentRow, 2)%>"  style='width:100%;text-align:center' onblur="t1chg(<%=CurrentRow-1%>)" maxlength=5>
						</TD>
						<TD ALIGN=CENTER>
							<input type="text" name="T2"  value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>"  style='width:100%;text-align:center' onblur="t2chg(<%=CurrentRow-1%>)"  maxlength=5>
						</TD>
						<TD ALIGN=center>		
							<input type="text" name="Toth" class=readonly8 value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>"  style="width:100%;text-align:right;height:22" readonly  >
						</TD>
						<TD ALIGN=CENTER>	
							<input type="text" name=kzhour   value="<%=tmpRec(CurrentPage, CurrentRow, 10)%>" style='width:100%;text-align:right;<%if tmpRec(CurrentPage, CurrentRow, 10)="0" then%>color:#red<%else%>color:red<%end if%>' onchange="datachg(<%=CurrentRow-1%>)">
						</TD>
						<TD ALIGN=CENTER>
							<input type="text" name=forget   value="<%=tmpRec(CurrentPage, CurrentRow, 9)%>" style='width:100%;text-align:right;<%if tmpRec(CurrentPage, CurrentRow, 9)="0" then%>color:#blue<%else%>color:blue<%end if%>' onchange="datachg(<%=CurrentRow-1%>)" > 
						</TD>
						<TD ALIGN=CENTER><!--遲到-->		
							<input type="text" name=latefor  value="<%=tmpRec(CurrentPage, CurrentRow, 12)%>" style='width:100%;text-align:right;<%if tmpRec(CurrentPage, CurrentRow, 12)="0" then%>color:#blue<%else%>color:blue<%end if%>' onchange="datachg(<%=CurrentRow-1%>)">
						</TD>
						<TD ALIGN=CENTER>
							<input type="text" name=JB  value="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" style='width:100%;text-align:right;<%if tmpRec(CurrentPage, CurrentRow, 25)="0" then%>color:#blue<%else%>color:blue<%end if%>' onchange="datachg(<%=CurrentRow-1%>)">
						</TD>
						<TD ALIGN=CENTER>			
							<input type="text" name=B3   value="<%=tmpRec(CurrentPage, CurrentRow, 16)%>" style='width:100%;text-align:right;<%if tmpRec(CurrentPage, CurrentRow, 16)="0" then%>color:#blue<%else%>color:blue<%end if%>' onchange="datachg(<%=CurrentRow-1%>)">
						</TD>		
						<TD ALIGN=CENTER>			
							<input type="text" name=lsempid  class=readonly8 readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" style="width:100%">
						</TD>		
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 17)%>
						</TD>
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 18)%>
						</TD>
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 19)%>
						</TD>
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 20)%>
						</TD>
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 21)%>
						</TD>
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 22)%>
						</TD>
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 23)%>
						</TD>
						<TD ALIGN=CENTER bgcolor="#FBE5CE">
							<%=tmpRec(CurrentPage, CurrentRow, 24)%>
						</TD>								 
					</TR> 

					<%else%>
						<input type=hidden name="empid">
						<input type=hidden name="workdat">
						<input type=hidden name="T1">
						<input type=hidden name="T2">
						<input type=hidden name="toth" value="0">
						<input type=hidden name="kzhour" value="0">
						<input type=hidden name="forget" value="0">
						<input type=hidden name="latefor" value="0" >
						<input type=hidden name="JB" value="0">
						<input type=hidden name="B3" value="0">		
						<input type=hidden name="lsempid" value="">		
						<input type=hidden name="op" value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>">
						<input type=hidden name="TotJia" value="0">
						<input type=hidden name="dsts" value="">
						<input type=hidden name="BC" value="">
					<%end if %>
					<%next%>
					<input type=hidden name="empid">
					<input type=hidden name="workdat">
					<input type=hidden name="T1">
					<input type=hidden name="T2">
					<input type=hidden name="toth">
					<input type=hidden name="kzhour" value="0">
					<input type=hidden name="forget" value="0">
					<input type=hidden name="latefor" value="0" >
					<input type=hidden name="JB" value="0">
					<input type=hidden name="B3" value="0">	
					<input type=hidden name="lsempid" value="">		
					<input type=hidden name="op" value="">
					<input type=hidden name="TotJia" value="0">
					<input type=hidden name="dsts" value="">
					<input type=hidden name="BC" value="">
				</TABLE>
			</td>
		</tr>
		<tr>
			<td>
				<table class="table-borderless table-sm bg-white text-secondary">
					<tr class="font9">
						<td align="CENTER" height=40 >

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
						PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
						</td> 
						
						<td align="CENTER" height=40  >
							<input type=BUTTON name=send value="(Y)Confirm"  class="btn btn-sm btn-danger" onclick="go()" >
							<input type=BUTTON name=send value="(N)Cancel"  class="btn btn-sm btn-outline-secondary"  >
							<input type=BUTTON name=send value="(B)Back Main"  class="btn btn-sm btn-outline-secondary"  onclick="gob()" >
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
	tmpRec = Session("YEDE02B")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("T1")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("T2")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 7) = request("toth")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 9) = request("forget")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 10) = request("kzhour")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 12) = request("lateFor")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 16) = request("B3")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 25) = request("JB")(CurrentRow) 
		'response.write CurrentPage & "-" & CurrentRow &"-" &  tmpRec(CurrentPage, CurrentRow, 0) & "-"& tmpRec(CurrentPage, CurrentRow, 3) &"<BR>"
	next
	Session("YEDE02B") = tmpRec
End Sub
%> 
<script language=vbscript >
function gob()
	open "<%=self%>.asp", "_self"
end function    
function go()
	'alert "ok" 
	if <%=self%>.UID.value="LSARY" then 
		if ( cdbl(<%=self%>.sum_h1.value)+cdbl(<%=self%>.sum_h2.value) )+ cdbl(<%=self%>.totjiaH.value) > 300  then 
			alert "加班時數超過(>)300小時,不可再加班!!"
			exit function
		else
			<%=self%>.action="<%=self%>.upd.asp"
			<%=self%>.submit()		
		end if	
	else
		<%=self%>.action="<%=self%>.upd.asp"
		<%=self%>.submit()
	end if	
end function 

function datachg(index)
	<%=self%>.op(index).value="*"
	databack(index)
end function 


function showWorkTime(index) 
	empidstr = <%=self%>.empid(index).value 	
	workdatstr = <%=self%>.workdat(index).value 
	yymmstr = left(replace(<%=self%>.workdat(index).value ,"/",""),6)
	'alert empidstr
	'alert workdatstr
	open "showWorkTime.asp?empid=" & empidstr &"&workdat=" & workdatstr &"&yymm=" & yymmstr   , "_blank"   , "top=100, left=100, width=500, height=400, scrollbars=yes  " 
end function  

function showTime(index) 
	empidstr = <%=self%>.empid(index).value 	
	workdatstr = <%=self%>.workdat(index).value 
	yymmstr = left(replace(<%=self%>.workdat(index).value ,"/",""),6)
	'alert empidstr
	'alert workdatstr
	open "showTime.asp?empid=" & empidstr &"&workdat=" & workdatstr &"&yymm=" & yymmstr   , "_blank"   , "top=100, left=100, width=350, height=400, scrollbars=yes  " 
end function  

function t1chg(index)
	if <%=self%>.t1(index).value<>"" then 
		CT1=left(trim(<%=self%>.t1(index).value),2)&":"&right(trim(<%=self%>.t1(index).value),2)
		<%=self%>.t1(index).value=CT1
		<%=self%>.op(index).value="*"		
		calctoth(index)		
	end if 
end function 
function t2chg(index)
	if <%=self%>.t2(index).value<>"" then 
		CT2=left(trim(<%=self%>.t2(index).value),2)&":"&right(trim(<%=self%>.t2(index).value),2)
		<%=self%>.t2(index).value=CT2
		<%=self%>.op(index).value="*"
		calctoth(index)
	end if 
end function 

function calctoth(index) 
	Cdat=trim(<%=self%>.workdat(index).value)
	CTime1=<%=self%>.t1(index).value
	CTime2=<%=self%>.t2(index).value
	'alert Cdat 
	'alert CTime1 
	'alert CTime2 
	DD1 = Cdat&" "&CTime1
	DD2 = Cdat&" "&CTime2	
	if CTime1<>"" and CTime2<>"" then 
		'工時
		if CTime2 >= CTime1 then 
			time_tot=fix( DATEDIFF("N", DD1, DD2 ) /30 )*0.5 
			'alert time_tot		
		else
			time_tot=(fix(DATEDIFF("N", DD1, DD2)/30)-1)*0.5 + 24 
			'alert time_tot	
		end if  
		<%=self%>.totH(index).value=time_tot 
		'曠職
		if time_tot < 8 and <%=self%>.dsts(index).value="H1"  then 
			<%=self%>.kzhour(index).value = 8- cdbl(<%=self%>.totH(index).value)-cdbl(<%=self%>.totjia(index).value)
		else
			<%=self%>.kzhour(index).value = 0
		end if 	  
		'忘刷卡
		if (<%=self%>.t1(index).value<>"" and <%=self%>.t2(index).value<>"") and  CTime1=CTime2 and <%=self%>.dsts(index).value="H1"  then 
			<%=self%>.forget(index).value="1"
		else
			<%=self%>.forget(index).value = "0"	
		end if  
		'加班
		if time_tot > 8 then  
			if <%=self%>.dsts(index).value="H1"  then 
				if left(<%=self%>.BC(index).value,1)="A" or right(trim(<%=self%>.BC(index).value),1)="D" then 
					<%=self%>.JB(index).value = time_tot - 8
					<%=self%>.JB(index).style.color="blue"
				else
					<%=self%>.JB(index).value = time_tot - 9
					<%=self%>.JB(index).style.color="blue"
				end if 	
			else
				if left(<%=self%>.BC(index).value,1)="A" then 
					<%=self%>.JB(index).value = time_tot 
				else
					<%=self%>.JB(index).value = time_tot-1 
				end if 	
				<%=self%>.JB(index).style.color="blue"
			end if 
		elseif time_tot= 0 then
			<%=self%>.JB(index).value = 0 
			<%=self%>.JB(index).style.color="blue"
			<%=self%>.B3(index).value = 0 
			<%=self%>.B3(index).style.color="blue"
		end if 	
		'夜班  (自21:00開始算  2009/09/15 by elin ) 
		if CTime2 < CTime1  then 
			BD1 = Cdat&" "&"21:00"
			B3_time=(fix(DATEDIFF("N", BD1, DD2)/30)-1)*0.5 + 24 
			<%=self%>.B3(index).value = B3_time 
			<%=self%>.B3(index).style.color="blue"
		else
			<%=self%>.B3(index).value = 0 
			<%=self%>.B3(index).style.color="blue"	
		end if 
	end if 	 
	databack(index)
end function 

function databack(index)
	C_T1=<%=self%>.t1(index).value
	C_T2=<%=self%>.t2(index).value
	C_forget=<%=self%>.forget(index).value
	C_kzhour=<%=self%>.kzhour(index).value
	C_latefor=<%=self%>.latefor(index).value
	C_jb=<%=self%>.jb(index).value
	C_b3=<%=self%>.b3(index).value
	C_toth=<%=self%>.toth(index).value
	
	open "<%=self%>.back.asp?func=upd&index="& index &"&CurrentPage="& <%=CurrentPage%> &_
		 "&code01=" & C_T1 & "&code02="& C_T2 &"&code03=" & C_forget &"&code04=" & C_kzhour &_
		 "&code05=" &C_latefor&"&code06=" & C_JB &"&code07="&C_B3 &"&code08=" & C_toth , "Back" 		 
	'parent.best.cols="50%,50%"	 
end function 


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.D1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.D2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.D1.value=ANS
		elseif a=2 then
			Document.<%=self%>.D2.value=ANS
			gos()	
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.D1.value=""
			Document.<%=self%>.D1.focus()
		elseif a=2 then
			Document.<%=self%>.D2.value=""
			Document.<%=self%>.D2.focus() 
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION 
</script>


