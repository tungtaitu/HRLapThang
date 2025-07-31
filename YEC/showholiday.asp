<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "showholiday"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
Set rsA = Server.CreateObject("ADODB.Recordset")   
Set rsB = Server.CreateObject("ADODB.Recordset")   
Set rsC = Server.CreateObject("ADODB.Recordset")   

'DAT1 = REQUEST("DAT1")
'DAT2 = REQUEST("DAT2")
yymm=trim(request("yymm"))
'whsno = trim(request("whsno"))
'unitno = trim(request("unitno"))
'groupid = trim(request("groupid"))
'zuno = trim(request("zuno"))
'job = trim(request("job")) 
'country = trim(request("country"))  
QUERYX = trim(request("empid"))   

if right(yymm,2)="12" then
	ccdt = cstr(left(YYMM,4)+1)&"/01/01"    '下個月第1天
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"   '下個月第1天 
end if 

nowmonth = year(date())&right("00"&month(date()),2)    '本月 
calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"    '本月第1天
 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS)
ENDdat=year(ENDdat)&"/"&right("00"&month(Enddat),2)&"/"&right("00"&day(Enddat),2) 

year_tx = 0 
gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record  

SQL="SELECT  A.JIATYPE,  CONVERT(CHAR(10), A.DATEUP, 111) DATEUP , A.TIMEUP, convert(char(10) , A.DATEDOWN , 111) datedown, "
SQL=SQL&"A.TIMEDOWN , A.HHOUR, A.MEMO AS JIAMEMO  , a.autoid as jiaid, isnull(a.xjsts,'') xjsts , x.sys_value as jbname,  B.*   FROM   "
SQL=SQL&"( SELECT * FROM EMPHOLIDAY where empid='"& QUERYX &"'  ) A  "
SQL=SQL&"LEFT JOIN ( SELECT * FROM view_empfile ) B ON B.EMPID = A.EMPID  	 "
SQL=SQL&"left join (select * from basicCode where func='JB' ) x on x.sys_type = a.jiatype "  
SQL=SQL&"order by b.empid, A.DATEUP desc "  
'response.write sql 
'RESPONSE.END  
sql1="select b.sys_value as jbname , a.* from "&_
				 "(select empid, jiatype, sum(hhour) hhour from empholiday where  empid='"& QUERYX &"' group by empid, jiatype ) a  "&_
				 "left join (select * from basicCode where func='JB' ) b on b.sys_type = a.jiatype "
		rsA.open sql1, conn, 1, 3 
		
		sql2="select b.sys_value as jbname , a.* from "&_
				 "(select empid, jiatype, sum(hhour) hhour from empholiday where  convert(char(6), dateup, 112)='"& yymm &"' and empid='"& QUERYX &"' group by empid, jiatype  ) a "	&_
				 "left join (select * from basicCode where func='JB' ) b on b.sys_type = a.jiatype "
		rsB.open sql2, conn, 1, 3 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 1 
	IF NOT RS.EOF THEN 
		chkemp = rs("empid")
		PageRec = rs.RecordCount 
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
		
		 
		
		'response.write rs("outdate")
		if trim(rs("outdate"))="" then 
			tx_enddat = ENDdat 			
		else			
			if right(rs("outdate"),5)>= "03/31"  and  year(rs("indat")) < year(date()) then
				'tx_enddat = cstr(left(yymm,4))&"/04/01"
				tx_enddat = rs("outdate") 
			else
				tx_enddat = rs("outdate") 
			end if			
		end if 	
		
		if rs("country")="VN" then 
			if  right(yymm,2)<="03" and yymm<=nowmonth then 
				if cdate( rs("calcTxdat") ) <= cdate(cstr(left(yymm,4)-1)&"/03/31") then 
					cc_indat = cstr(left(yymm,4)-1)&"/04/01" 
				else
					cc_indat = rs("calcTxdat")
				end if	
			else
				cc_indat = cstr(left(yymm,4))&"/04/01" 		
			end if  
		else
			cc_indat= cstr(left(yymm,4))&"/01/01" 		
		end if 
		
		if rs("country")="VN" then 
			sql3="select empid, jiatype, sum(hhour) hhour from empholiday  "&_
				 "where jiatype='E' and convert(char(10), dateup, 111) between '"& cc_indat &"' "&_
				 "and '"& tx_enddat &"'  and empid='"& QUERYX &"' group by empid, jiatype  "		
		else
			sql3="select empid, jiatype, sum(hhour) hhour from empholiday  "&_
				 "where jiatype='I' and convert(char(10), dateup, 111) between '"& cc_indat &"' "&_
				 "and '"& tx_enddat &"'  and empid='"& QUERYX &"' group by empid, jiatype  "				
		end if
		rsC.open sql3, conn, 1, 3 	
		
		'response.write sql3 
		'response.end 			
		year_tx = datediff("m", cdate(cc_indat) , cdate(tx_enddat) )
		
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
				tmpRec(i, j, 17)=RS("DATEUP")
				tmpRec(i, j, 18)=RS("TIMEUP")
				tmpRec(i, j, 19)=RS("DATEDOWN")
				tmpRec(i, j, 20)=RS("TIMEDOWN")
				tmpRec(i, j, 21)=RS("JIAMEMO")
				tmpRec(i, j, 22)=RS("JIATYPE") 
				tmpRec(i, j, 23)=RS("hhour") 
				tmpRec(i, j, 24)=RS("jiaID") 
				tmpRec(i, j, 25)=RS("tx") 
				tmpRec(i, j, 26)=RS("outdate") 
				tmpRec(i, j, 27)=RS("jbname") 
				tmpRec(i, j, 28)=RS("xjsts") 
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
	Session("showholidayB") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("showholidayB")
	
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


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2) 
	
END FUNCTION 

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12" 
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	  



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
	'<%=self%>.QUERYX.focus()	
end function   

function datachg()
	<%=SELF%>.totalpage.VALUE=0
	<%=self%>.action="<%=SELF%>.fore.asp?totalpage=0"
	<%=self%>.submit
end function 

-->
</SCRIPT>  
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=SELF%>.fore.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	 
<INPUT NAME=DAT1 VALUE="<%=DAT1%>" TYPE=HIDDEN >
<INPUT NAME=DAT2 VALUE="<%=DAT2%>" TYPE=HIDDEN  >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	員工請假查詢 TRA PHÉP NHÂN VIÊN </TD></tr>
</table> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<!-------------------------------------------------------------------->  
<table width=500 class=txt  cellspacing="1" cellpadding="1" BGCOLOR="#E4E4E4" >
	<tr bgcolor=#ffffff>
		<td width=70 align=right>Mã số: </td>
		<td width=60 ><%=tmpRec(CurrentPage,index + 1,1)%> </td>
		<td width=60  align=right>Họ tên:</td>
		<td colspan=3><%=tmpRec(CurrentPage,index + 1,2)%>&nbsp;<%=tmpRec(CurrentPage,index + 1,3)%> </td>
	</tr>
	<tr  bgcolor=#ffffff>	
		<td  align=right>到職日<br>Ngày nhập xưởng: </td>
		<td   ><%=tmpRec(CurrentPage,index + 1,5)%></td>
		<td align=right>年假<br>Phép năm:</td>
		<td><%=tmpRec(CurrentPage,index + 1,25)%>=<%=cdbl(tmpRec(CurrentPage,index + 1,25))*8%>H</td>
		<td align=right>離職日<br>Ngày thôi việc:</td>
		<td width=80 nowrap  ><%=tmpRec(CurrentPage,index + 1,26)%></td>
	</tr>  	
</table>
<br>
<table width=500  class=txt9 bgcolor=#e4e4e4 cellspacing="1" cellpadding="1" > 
	<tr  bgcolor=MistyRose>
		<td colspan=2>------------所有假別總計 THÔNG KÊ TẤT CẢ PHÉP--------------</td>
	</tr>
	<tr bgcolor=lightyellow>
		<td>假別 Phép: </td>
		<td>總時數 Tổng số giờ:</td>
	</tr>
	<%if  chkemp<>"" then %>
	<%while not rsa.eof  
	%>
	<tr bgcolor=#ffffff>
		<td><%=rsa("jbname")%></td>
		<td><%=rsa("hhour")%></td>
	</tr>
	<%
	rsa.movenext
	wend 
	set rsa=nothing 
	jiastr ="" 
	%>	
	<%end if%>
	<tr  bgcolor=MistyRose>
		<%if tmpRec(CurrentPage,index + 1,4)="VN" then %>
			<td colspan=2>------------年度年假統計 THỐNG KÊ PHÉP NĂM--------------(<%=cdbl(year_tx)*8%>)小時 Giờ &nbsp; <%=cc_indat%>~<%=tx_enddat%></td>
		<%else%>
			<td colspan=2>------------年度返鄉休假統計 THÔNG KÊ PHÉP HỒI HƯƠNG--------------(<%=cdbl(year_tx)*8%>)小時 Giờ &nbsp; <%=cc_indat%>~<%=tx_enddat%></td>
		<%end if%>
	</tr>
	<%if   chkemp<>"" then %>
	<%while not rsC.eof    
	%>
	<tr bgcolor=#ffffff>
		<td><%if tmpRec(CurrentPage,index + 1,4)="VN" then %>(E) 年假 PHÉP NĂM<%else%>(I) 返鄉休假 PHÉP HỒI HƯƠNG <%end if%></td>
		<td><%if tmpRec(CurrentPage,index + 1,4)="VN" then %><%=rsC("hhour")%><%else%><%=rsC("hhour")%>,共<%=cdbl(rsC("hhour"))/8%>天<%end if%></td>
	</tr>
	<%
	rsC.movenext
	wend 
	set rsC=nothing 
	%> 			
	<%end if%>
	<tr  bgcolor=MistyRose>
		<td colspan=2>------------本月請假統計 THỐNG PHÉP TRONG THÁNG--------------<%=yymm%></td>
	</tr>	
	<%if chkemp<>"" then %>
	<%while not rsb.eof 	 
	%>	
	<tr bgcolor=#ffffff>
		<td><%=rsb("jbname")%></td>
		<td><%=rsb("hhour")%></td>
	</tr>
	<%
	rsb.movenext
	wend 
	set rsb=nothing 
	jiastr ="" 
	%>	
	<%end if%>		
</table> 
<table width=550 class=txt9 cellpadding="0" >
	<tr BGCOLOR="LightGrey" height=22>		
 		<TD align=center  >STT</TD>
 		<TD align=center  >假別<br>Phép</TD>
 		<TD width=80 align=center nowrap >日期(起)<br>Từ ngày</TD>
		<TD align=center  >時間(起)<br>Từ giờ</TD>
		<TD width=80 align=center nowrap >日期(迄)<br>Đến ngày</TD>
		<td align=center  >時間(迄)<br>Tới giờ</td>
		<td align=center  >時數<br>Số giờ</td>
		<td align=center >事由<br>Lý do</td>		
	</tr>
	 
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR=""
		END IF 	 
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
	%>
	<TR BGCOLOR=<%=WKCOLOR%> > 	   
		<Td class=txt8><%=CurrentRow%></td>
 		<TD>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN  %>	 				
	 			<INPUT TYPE=HIDDEN NAME=HOLIDAY_TYPE  >
	 			<INPUT NAME=HOLIDASTR value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>-<%=tmpRec(CurrentPage, CurrentRow, 27)%>" class=readonly8  readonly size=8  > 	 			 
			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HOLIDAY_TYPE  >	
				<INPUT TYPE=HIDDEN NAME=HOLIDASTR >			
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=HHDAT1 size=12 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>" > 				
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HHDAT1  >								
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=HHTIM1 size=7 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HHTIM1  >				
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	 				
 				<input name=HHDAT2 size=12 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 19)%>" >
 			<%ELSE%>					
				<INPUT TYPE=HIDDEN NAME=HHDAT2  >			
			<%END IF %>
 		</TD> 
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=HHTIM2 size=7 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HHTIM2  >				
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=toth size=4 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" style="text-align:right">
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=toth  >				
			<%END IF %>
 		</TD> 
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN  
				reson = tmpRec(CurrentPage, CurrentRow, 21) 
				if tmpRec(CurrentPage, CurrentRow, 28)="C" then reson="不扣全勤,"&reson
			%>	
 				<input name=JIAMEMO size=18 class=readonly readonly   value="<%=reson%>" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=JIAMEMO  >				
			<%END IF %>
 		</TD> 		
	</TR>
	<%next%>  
</table>	

 
<TABLE border=0 width=600 class=font9 >
<tr> 
	<td align=center>
		<input type="BUTTON" name="send" value="關閉視窗(Close)Đóng" class=button onclick="vbscript:window.close()" >		
	</td>	 
</TR>
</TABLE>  
<input type=hidden name=func >
<input type=hidden name=op >
<input type=hidden name=empid >

</form>

</body>
</html>

<script language=vbscript> 
function del(index) 
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="del" 		
		open "<%=self%>.back.asp?func=del&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
	else
		<%=self%>.op(index).value="no"  
		open "<%=self%>.back.asp?func=no&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
	end if 	 	
	'parent.best.cols="70%,30%"
end function 

function BACKMAIN()	
	open "../main.asp" , "_self"
end function   

function oktest(N)	
	tp=<%=self%>.totalpage.value 
	cp=<%=self%>.CurrentPage.value 
	rc=<%=self%>.RecordInDB.value 
	'open "empworkB.fore.asp?empautoid="& N &"&yymm="&"<%=calcmonth%>", "_self" 
	open "empworkB.fore.asp?empautoid="& N &"&YYMM="&"<%=calcmonth%>" &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_self" 
end function   

FUNCTION CLR()
	OPEN "<%=SELF%>.ASP" , "_self"
END FUNCTION 

function strchg(a)
	if a=1 then 
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value) 
		'IF TRIM(<%=self%>.empid.value)<>"" THEN 
			<%=SELF%>.totalpage.VALUE=0
			<%=SELF%>.ACTION="<%=SELF%>.FORE.ASP?TOTALPAE=0"
			<%=SELF%>.SUBMIT()
		'END IF 
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function   

function go()
	<%=self%>.action="<%=self%>.updateDB.asp" 
	<%=self%>.submit()
end function 
	
</script>

