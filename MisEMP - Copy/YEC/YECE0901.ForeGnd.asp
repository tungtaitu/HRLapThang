<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next
session.codepage="65001"
SELF = "YECE0901"

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

PERAGE = REQUEST("PERAGE")

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
MMDAYS = CDBL(days) 
'RESPONSE.WRITE  MMDAYS
'RESPONSE.END
'---------------------------------------------------------------------------------------- 

recalc  = request("recalc")
if recalc="Y" then 
	sql="delete empdsalary where yymm='"& YYMM &"' and isnull(country,'')='"& COUNTRY &"' and isnull(whsno,'') like '%"& whsno &"'"
	conn.execute(Sql)
end if 

sqlstr = "update empwork set kzhour=0 where yymm='"& YYMM &"'  and kzhour<0 "
conn.execute(sqlstr)

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 30    'number of fields per record

 	sqlx="select * from empdsalary_bak where country='"& country &"' and yymm='"& yymm &"'"
 	set rds=conn.execute(Sqlx)
 	if rds.eof then 
 		sql="select b.country countrydesc,b.empnam_vn, b.empnam_cn, gx.lwstr whsnodesc, gx.lgstr groupdesc, "&_
 			"jx.ljstr jobdesc, convert(char(10),b.indat,111) date1 , "&_
 			"gx.lz zuno, convert(char(10),b.outdat,111) as outdate , b.bhdat, a.*, isnull(a.flag,'') as updflag  , isnull(mz.money1,0) money1 from "&_
 			"(select * from empdsalary where yymm='"& yymm &"' and real_total > 0   ) a "&_
 			"join(select * from empfile ) b on b.empid = a.empid "&_
			"left join(select * from view_empgroup  where yymm='"& yymm &"' ) gx on gx.empid = a.empid 	"&_
			"left join(select * from view_empjob  where yymm='"& yymm &"' ) jx on jx.empid = a.empid 	"&_
			"left join ( select * from empmoney where yymm2='"& yymm &"' ) mz on mz.empid2=a.empid "&_
			"where b.empid<>'' " 
 	else
 		sql="select b.country countrydesc,b.empnam_vn, b.empnam_cn, gx.lwstr whsnodesc, gx.lgstr groupdesc, "&_
 			"jx.ljstr jobdesc, convert(char(10),b.indat,111) date1 , "&_
 			"gx.lz zuno, convert(char(10),b.outdat,111) as outdate , b.bhdat, a.*, isnull(a.flag,'') as updflag  , isnull(mz.money1,0) money1 from "&_
 			"(select * from empdsalary where yymm='"& yymm &"' and real_total > 0   ) a "&_
 			"join(select * from empfile ) b on b.empid = a.empid "&_ 
			"left join(select * from view_empgroup  where yymm='"& yymm &"' ) gx on gx.empid = a.empid 	"&_
			"left join(select * from view_empjob  where yymm='"& yymm &"' ) jx on jx.empid = a.empid 	"&_
			"left join ( select * from empmoney where yymm2='"& yymm &"' ) mz on mz.empid2=a.empid "&_
			"where b.empid<>'' " 
 	end if 
 	sql=sql&"and CONVERT(CHAR(10), b.indat, 111)< '"& ccdt &"' and ( isnull(b.outdat,'')='' or convert(char(10),b.outdat,111)>'"& calcdt &"' )  "&_
			"and a.whsno like '%"& whsno &"%' and a.groupid like '%"& groupid &"%'  "&_
			"and a.COUNTRY like '%"& COUNTRY  &"%' and a.job like '%"& job &"%' and a.empid like '%"& QUERYX &"%' "
	if outemp="D" then
		sql=sql&" and ( isnull(ltrim(rtrim(b.outdate)),'')<>'' and  ltrim(rtrim(b.outdate))>'"& calcdt &"' )  "
	end if
	if outemp="B" then
		sql=sql&" and ltrim(rtrim(isnull(b.bankid,'')))=''  "
	end if

sql=sql&"order by a.whsno, a.groupid, a.empid   "
'response.write sql
'response.end
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open SQL, conn, 3, 3
	IF NOT RS.EOF THEN
		PageRec =  rs.RecordCount
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF
	'response.write TotalPage 
	'response.write PageRec 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("date1")  '到職日期
				tmpRec(i, j, 6) = rs("job")
				tmpRec(i, j, 7) = rs("whsno")
				tmpRec(i, j, 8) = "" 'rs("unitno")
				tmpRec(i, j, 9)	=RS("groupid")
				tmpRec(i, j, 10)=RS("zuno")
				tmpRec(i, j, 11)=RS("whsnodesc")
				tmpRec(i, j, 12)="" 'RS("unitdesc")
				tmpRec(i, j, 13)=RS("groupdesc")
				tmpRec(i, j, 14)="" 'RS("zunodesc")
				tmpRec(i, j, 15)=RS("jobdesc")
				tmpRec(i, j, 16)=RS("countrydesc")
				tmpRec(i, j, 17)="" 'RS("autoid")
				IF RS("zuno")="XX" THEN
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("zuno")
				END IF
				tmpRec(i, j, 19) = rs("workdays")
				if rs("country")="VN" then 
					tmpRec(i, j, 20) = rs("laonh")
				else
					tmpRec(i, j, 20) = rs("real_total")
				end if 	
				tmpRec(i, j, 21) = rs("outdate")
				tmpRec(i, j, 22) = rs("dm") 
				if rs("updflag")=""  then   '現金
					if  ( rs("country")="VN" and yymm>="200804" ) and rs("groupid")<>"A051"  then 
						tmpRec(i, j, 23) = cdbl(rs("JX")) + cdbl(rs("money1"))
					else
						tmpRec(i, j, 23) = cdbl(rs("xianM")) 
					end if 	
				else
					tmpRec(i, j, 23) = cdbl(rs("xianM")) 
				end if 	
				tmpRec(i, j, 23) = cdbl(rs("xianM"))  
				if trim(rs("outdate"))<>"" and (rs("xianM")) = "0" then 
					tmpRec(i, j, 23) = rs("laonh")
					tmpRec(i, j, 27) = "red"
				else	
					tmpRec(i, j, 27) = "blue"
				end if 	
				'轉帳				
				tmpRec(i, j, 24) = cdbl(rs("laonh"))-cdbl(tmpRec(i, j, 23))-cdbl(rs("dkm"))   'cdbl(rs("zhuanM"))-cdbl(rs("dkm"))-cdbl(rs("jx"))
				
				tmpRec(i, j, 25) = rs("dkm") 
				tmpRec(i, j, 26) = rs("JX") 				
				tmpRec(i, j, 27) = rs("money1") 	
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
	Session("YECE0901F") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YECE0901F")
	tot1 = 0
	tot2 = 0
	tot3 = 0
	COUNTRY = REQUEST("COUNTRY")

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

'response.end
FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)

END FUNCTION 
'Response.ContentType = "application/vnd.ms-excel" 
'Response.ContentType = "application/vnd.ms-excel;charset=big5"  
'Response.AppendHeader("content-disposition", "attachment; filename=aaa.xls")
%>

<html>

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
 
</head>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
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
</SCRIPT> 
<body   topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"   bgproperties="fixed"  >
<form name="<%=self%>" method="post" action="<%=SELF%>.ForeGnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>">
<INPUT TYPE=hidden NAME=MMDAYS VALUE="<%=MMDAYS%>">
<INPUT TYPE=hidden NAME=COUNTRY VALUE="<%=COUNTRY%>">  

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<TD>計薪年月Thống kê tiền lương：<%=YYMM%></TD>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%">
								<TR HEIGHT=25 BGCOLOR="LightGrey" CLASS="txt8"  >
									<TD align=center width=40 nowrap>國籍<br>Quốc tịch</TD> 		 
									<TD align=center width=40 nowrap>廠別<br>Xưởng</TD> 		 
									<TD align=center nowrap>單位<br>Bộ phận</TD>
									<TD align=center width=45 nowrap>工號<br>Mã số NV</TD> 		
									<TD width=110 nowrap >員工姓名(中,英,越)<br>Tên nhân viên</TD>
									<td align=center width=60 nowrap>到職日期<br>Ngày nhập xưởng</td>
									<td align=center width=60 nowrap>離職日期<br>Ngày thôi việc</td> 			
									<TD align=center width=60 >本月薪資<br>Tiền lương trong tháng</TD>
									<TD align=center width=50 >績效獎金<br>Tiền thưởng hiệu xuất</TD>
									<TD align=center width=50 >暫扣款<br>Tạm tính </TD>
									<TD align=center width=50 >特別獎金<br>Tiên thưởng đặc biệt</TD>
									<TD align=center width=40 nowrap>幣別<br>Tiền tệ</TD>
									<td align=center  nowrap>轉帳<br>Chuyển <br>khoản</td>
									<td align=center  nowrap>現金<br>Tiền mặt</td>  		 
								</TR> 	
								<%	
								for CurrentRow = 1 to PageRec
									IF CurrentRow MOD 2 = 0 THEN
										WKCOLOR="LavenderBlush"
									ELSE
										WKCOLOR="#DFEFFF"
									END IF
									tot1 = tot1 + round(tmpRec(CurrentPage, CurrentRow, 20),0) 
									tot2 = tot2 + round(tmpRec(CurrentPage, CurrentRow, 23),0)  '現金加總
									tot3 = tot3 + round(tmpRec(CurrentPage, CurrentRow, 24),0)  '轉帳加總
									tot4 = tot4 + round(tmpRec(CurrentPage, CurrentRow, 26),0) 
									tot5 = tot5 + round(tmpRec(CurrentPage, CurrentRow, 27),0) 
									'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
								%>
								<TR BGCOLOR=<%=WKCOLOR%> height="20" CLASS="txt8">
									<TD nowrap align=center><%=tmpRec(CurrentPage, CurrentRow, 16)%></TD>		
									<TD nowrap align=center ><%=tmpRec(CurrentPage, CurrentRow, 11)%></TD>
									<TD nowrap align=center ><%=tmpRec(CurrentPage, CurrentRow, 13)%></TD>
									<TD align=center nowrap><%=tmpRec(CurrentPage, CurrentRow, 1)%></TD>
									<TD  nowrap><%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;<%=left(tmpRec(CurrentPage, CurrentRow, 3),10)%> </TD>
									<TD align=center nowrap><%=tmpRec(CurrentPage, CurrentRow, 5)%></TD>
									<TD align=center nowrap><%=tmpRec(CurrentPage, CurrentRow, 21)%></TD>
									<TD  align=right nowrap><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 20),0)%></TD>
									<TD  align=right><%=tmpRec(CurrentPage, CurrentRow, 26)%></TD>
									<TD  align=right><%=tmpRec(CurrentPage, CurrentRow, 25)%></TD>
									<TD  align=right><%=tmpRec(CurrentPage, CurrentRow, 27)%></TD><!--特別獎金-->
									<TD  align=center><%=tmpRec(CurrentPage, CurrentRow, 22)%></TD>
									<TD   >
									<%if tmpRec(CurrentPage, CurrentRow, 1)="" then %>
										<INPUT TYPE=HIDDEN NAME=ZHUANM >
									<%else%>
										<INPUT type="text"  NAME=ZHUANM value="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" class="INPUTBOX8" STYLE="width:100%;TEXT-ALIGN:RIGHT" onblur="zhuanmchg(<%=currentRow-1%>)"></TD>
									<%end if%>	
									<INPUT TYPE=HIDDEN NAME=ZHUANMB  value="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" >
									<TD  >
									<%if tmpRec(CurrentPage, CurrentRow, 1)="" then %>
										<INPUT TYPE=HIDDEN NAME=XIANM   >
									<%else%>
										<INPUT type="text" NAME=XIANM value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>"  class="INPUTBOX8"  STYLE="width:100%;TEXT-ALIGN:RIGHT;color:<%=tmpRec(CurrentPage, CurrentRow, 27)%>"   onblur="xianMchg(<%=currentRow-1%>)"   ></TD>
									<%end if%>
									<INPUT TYPE=HIDDEN NAME=XIANMB  value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" >		
									<input name="RELTOTMONEY" type="hidden" value="<%=tmpRec(CurrentPage, CurrentRow, 20)%>">
								</TR>
								<%next%>
								<tr bgcolor="#E4E4e4" CLASS="txt8">
									<td colspan=7 >總計 Thống kê</td>
									<td align=right><input type="text" name=tot1 value="<%=tot1%>" class=INPUTBOX8  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue"  readonly   ></td>
									<td ><input type="text" name=tot4 value="<%=(tot4)%>" class=INPUTBOX8  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue"  readonly    ></td>
									<td ></td>
									<td ><input type="text" name=tot5 value="<%=(tot5)%>" class=INPUTBOX8  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue"  readonly   ></td>
									<td ></td>
									<td ><input type="text" name=tot2 value="<%=(tot3)%>" class=INPUTBOX8  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue"  readonly    ></td>
									<td ><input type="text" name=tot3 value="<%=(tot2)%>" class=INPUTBOX8  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue"  readonly    ></td>
								</tr>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<INPUT TYPE=HIDDEN NAME=XIANM   value="0">
							<INPUT TYPE=HIDDEN NAME=ZHUANM value="0">
							<input name="RELTOTMONEY" type="hidden" value="0"> 
							<INPUT TYPE=HIDDEN NAME=ZHUANMB value="0">
							<INPUT TYPE=HIDDEN NAME=XIANMB value="0">									

							<table class="txt" cellpadding=3 cellspacing=3>
								<tr class=font9>
									<td align="CENTER" height=40 WIDTH=75%>
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
									<FONT CLASS=TXT8>&nbsp;&nbsp;PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%></FONT>
									</TD>
									<TD WIDTH=25% ALIGN=RIGHT>
										<input type="BUTTON" name="send" value="確　認" class="btn btn-sm btn-danger" ONCLICK="GO()">
										<input type="BUTTON" name="send" value="取　消" class="btn btn-sm btn-outline-secondary" onclick="clr()">
									</TD>
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
	tmpRec = Session("YECE0901F")
	for CurrentRow = 1 to PageRec 
		tmpRec(CurrentPage, CurrentRow, 24) = request("ZHUANM")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("XIANM")(CurrentRow)
	next
	Session("YECE0901F") = tmpRec 
End Sub
%>

<script language=vbscript>
function BACKMAIN()
	open "../main.asp" , "_self"
end function

function clr()
	open "<%=SELF%>.fore.asp" , "_self"
end function

function go()	
	<%=self%>.action="<%=SELF%>.upd.asp"
	<%=self%>.submit()
end function
 


function zhuanmchg(index)
	'F_EXRT = CDBL(<%=SELF%>.EXRT(INDEX).VALUE)
	REL_TOTAMT = CDBL(<%=SELF%>.RELTOTMONEY(INDEX).VALUE)	
	yymmstr=<%=self%>.yymm.value 
	IF ISNUMERIC(<%=SELF%>.ZHUANM(INDEX).VALUE)=FALSE THEN 
		ALERT "請輸入數值 NHẬP SỐ!!" 
		<%=SELF%>.ZHUANM(INDEX).VALUE = <%=SELF%>.RELTOTMONEY(INDEX).VALUE 
		<%=SELF%>.XIANM(INDEX).VALUE = "0"
		'<%=SELF%>.ZHUANM(INDEX).SELECTED()
        <%=SELF%>.ZHUANM(INDEX).select()
		EXIT FUNCTION 
	END  IF 
	F_ZHUANM = CDBL(<%=SELF%>.ZHUANM(INDEX).VALUE)
	F_XIANM = CDBL(<%=SELF%>.XIANM(INDEX).VALUE)
	if  cdbl(F_ZHUANM)=CDBL(REL_TOTAMT)  then 
		<%=SELF%>.XIANM(INDEX).VALUE = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)		
	else 
		IF  CDBL(<%=SELF%>.ZHUANM(INDEX).VALUE)+cdbl(<%=SELF%>.XIANM(INDEX).VALUE)-CDBL(<%=SELF%>.ZHUANMB(INDEX).VALUE)  > CDBL(REL_TOTAMT)  THEN 
			ALERT "轉款金額輸入錯誤 NHẬP TIỀN LÃNH DÙNG SAI!!(大於實領金額 LỚN HƠN TIỀN LÃNH DÙNG)"
			<%=SELF%>.ZHUANM(INDEX).VALUE = <%=SELF%>.RELTOTMONEY(INDEX).VALUE 
			<%=SELF%>.XIANM(INDEX).VALUE = "0"
			'<%=SELF%>.ZHUANM(INDEX).SELECTED()
			<%=SELF%>.ZHUANM(INDEX).select()
			EXIT FUNCTION 
		END  IF  
	end if 	
	<%=SELF%>.XIANM(INDEX).VALUE = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)
    <%=SELF%>.XIANM(INDEX).select()        
    CODESTR01 = <%=SELF%>.ZHUANM(INDEX).VALUE  
    CODESTR02 = <%=SELF%>.XIANM(INDEX).VALUE
    T1 = <%=SELF%>.ZHUANMB(INDEX).VALUE  
    T2 = <%=SELF%>.ZHUANM(INDEX).VALUE  
    T3 = <%=SELF%>.XIANMB(INDEX).VALUE  
    T4 = <%=SELF%>.XIANM(INDEX).VALUE   
    
    <%=self%>.tot2.value =  cdbl(<%=self%>.tot2.value) - cdbl(t1)+cdbl(t2)
    <%=self%>.tot3.value =  cdbl(<%=self%>.tot3.value) - cdbl(t3)+cdbl(t4)
    open "<%=SELF%>.back.asp?ftype=ZXCHG&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01=" & CODESTR01 & "&CODESTR02=" &CODESTR02 & "&yymm="& yymmstr , "Back"
    'PARENT.BEST.COLS="70%,30%"
END FUNCTION    

function xianMchg(index)
	'F_EXRT = CDBL(<%=SELF%>.EXRT(INDEX).VALUE)
	REL_TOTAMT = CDBL(<%=SELF%>.RELTOTMONEY(INDEX).VALUE)	
	yymmstr=<%=self%>.yymm.value 
	IF ISNUMERIC(<%=SELF%>.XIANM(INDEX).VALUE)=FALSE THEN 
		ALERT "請輸入數值 NHẬP SAI SỐ!!" 
		'<%=SELF%>.ZHUANM(INDEX).VALUE = <%=SELF%>.RELTOTMONEY(INDEX).VALUE 
		<%=SELF%>.XIANM(INDEX).VALUE = "0"
		'<%=SELF%>.ZHUANM(INDEX).SELECTED()
        <%=SELF%>.ZHUANM(INDEX).select()
		EXIT FUNCTION 
	END  IF 
	F_ZHUANM = CDBL(<%=SELF%>.ZHUANM(INDEX).VALUE)
	F_XIANM = CDBL(<%=SELF%>.XIANM(INDEX).VALUE) 
	
	<%=SELF%>.XIANM(INDEX).VALUE = F_XIANM
	<%=SELF%>.ZHUANM(INDEX).VALUE = cdbl(REL_TOTAMT ) - cdbl(F_XIANM)
	
	'if  cdbl(F_ZHUANM)=CDBL(REL_TOTAMT)  then 
	'	<%=SELF%>.XIANM(INDEX).VALUE = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)		
	'else 
	'	IF  CDBL(<%=SELF%>.ZHUANM(INDEX).VALUE)+cdbl(<%=SELF%>.XIANM(INDEX).VALUE)-CDBL(<%=SELF%>.ZHUANMB(INDEX).VALUE)  > CDBL(REL_TOTAMT)  THEN 
	'		ALERT "轉款金額輸入錯誤 NHẬP TIỀN LÃNH DÙNG SAI!!(大於實領金額 LỚN HƠN TIỀN LÃNH DÙNG)"
	'		<%=SELF%>.ZHUANM(INDEX).VALUE = <%=SELF%>.RELTOTMONEY(INDEX).VALUE 
	'		<%=SELF%>.XIANM(INDEX).VALUE = "0"
	'		'<%=SELF%>.ZHUANM(INDEX).SELECTED()
	'		<%=SELF%>.ZHUANM(INDEX).select()
	'		EXIT FUNCTION 
	'	END  IF  
	'end if 	
	'<%=SELF%>.XIANM(INDEX).VALUE = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)
    '<%=SELF%>.XIANM(INDEX).FOCUS()        
    
    CODESTR01 = <%=SELF%>.ZHUANM(INDEX).VALUE  
    CODESTR02 = <%=SELF%>.XIANM(INDEX).VALUE
    T1 = <%=SELF%>.ZHUANMB(INDEX).VALUE  
    T2 = <%=SELF%>.ZHUANM(INDEX).VALUE  
    T3 = <%=SELF%>.XIANMB(INDEX).VALUE  
    T4 = <%=SELF%>.XIANM(INDEX).VALUE   
    
    <%=self%>.tot2.value =  cdbl(<%=self%>.tot2.value) - cdbl(t1)+cdbl(t2)
    <%=self%>.tot3.value =  cdbl(<%=self%>.tot3.value) - cdbl(t3)+cdbl(t4)
    open "<%=SELF%>.back.asp?ftype=ZXCHG&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01=" & CODESTR01 & "&CODESTR02=" &CODESTR02 & "&yymm="& yymmstr , "Back"
    'PARENT.BEST.COLS="70%,30%"
END FUNCTION   

 

function view1(index)
	yymmstr = <%=self%>.yymm.value
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	OPEN "../zzz/getempWorkTime.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  scrollbars=yes"
end function

</script>

