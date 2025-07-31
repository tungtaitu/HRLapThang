<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "showsalary"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

C_ym=request("yymm")

ym1 = left(c_ym,4)-1&right(c_ym,2)
ym2 = c_ym
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
empid1 = trim(REQUEST("empid1"))

ccnt = cdbl(ym2) - cdbl(ym1)
if ccnt = 0 then ccnt = 1 

gTotalPage = 1
PageRec = 20*ccnt   'number of records per page
TableRec = 80    'number of fields per record
'NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/01"
NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)
 
	sql="select isnull(e.lj,'') lj , isnull(e.ljstr,'') ljstr ,  "&_
		"isnull(d.lw,'') lw , isnull(d.lg,'') lg , isnull(d.lz,'') lz , isnull(d.ls,'') ls , "&_
		"isnull(d.lwstr,'') lwstr , isnull(d.lgstr,'') lgstr , isnull(d.lzstr,'') lzstr, isnull(d.lsstr,'') lsstr, "&_
		"c.empnam_cn, c.empnam_vn, c.nindat, c.outdate, a.* , isnull(a.real_total,0) as backTotal , f.sys_value as Sjstr , isnull(g.totamt,0) wpamt , c.taxcode  "&_
		"from "&_
		"( select * from empdsalary_bak where  yymm  between '"& ym1 &"' and '"&ym2&"' and empid like '%"&empid1&"' and country like '"&COUNTRY&"%' "&_
		"and whsno like '"&whsno&"%' and groupid like '"&groupid&"%'  ) a "&_		
		"join (select empid, empnam_cn, empnam_vn, convert(char(10),indat,111)  nindat , isnull(convert(char(10),outdat,111),'')  outdate, isnull(taxcode,'') taxcode from empfile ) c on c.empid = a.empid "&_
		"left join (select *from view_empgroup  ) d on d.empid = a.empid  and d.yymm = a.yymm   "&_
		"left join (select *from view_empjob) e on e.empid = a.empid and e.yymm = a.yymm "&_
		"left join (select *from basicCode) f on f.sys_type = a.job "&_
		"left join (select * from salarywp ) g on g.empid = a.empid and g.yymm = a.yymm "  
sql = sql & "order by a.empid , a.yymm " 
'response.write sql

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage 
		empname = trim(rs("empnam_cn")) &" "&trim(rs("empnam_vn")) 
		nindat = rs("nindat")
		country = rs("country")
		taxcode=rs("taxcode")
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
				tmpRec(i, j, 6) = rs("lj")
				tmpRec(i, j, 7) = rs("lw")
				tmpRec(i, j, 8) = rs("ls")
				tmpRec(i, j, 9)	=RS("lg")
				tmpRec(i, j, 10)=RS("lz")
				tmpRec(i, j, 11)=RS("lwstr")				
				tmpRec(i, j, 12)=RS("lgstr")
				tmpRec(i, j, 13)=RS("lzstr")
				tmpRec(i, j, 14)=RS("ljstr")
				tmpRec(i, j, 15)=RS("job")
				tmpRec(i, j, 16)=RS("whsno")
				tmpRec(i, j, 17)=RS("Sjstr") 				
				tmpRec(i, j, 18)=RS("BB")
				tmpRec(i, j, 19)=RS("CV")
				tmpRec(i, j, 20)=RS("PHU")
				tmpRec(i, j, 21)=RS("NN")
				tmpRec(i, j, 22)=RS("KT")
				tmpRec(i, j, 23)=RS("MT")
				tmpRec(i, j, 24)=RS("TTKH")
				tmpRec(i, j, 25)=RS("QC")
				tmpRec(i, j, 26)=RS("TNKH")
				tmpRec(i, j, 27)=RS("TBTR")
				tmpRec(i, j, 28)=RS("JX")
				tmpRec(i, j, 29)=RS("H1M")
				tmpRec(i, j, 30)=RS("H2M")
				tmpRec(i, j, 31)=RS("H3M")
				tmpRec(i, j, 32)=RS("B3M")
				tmpRec(i, j, 33)=RS("H1")
				tmpRec(i, j, 34)=RS("H2")
				tmpRec(i, j, 35)=RS("H3")
				tmpRec(i, j, 36)=RS("B3")
				tmpRec(i, j, 37)=RS("kzhour")
				tmpRec(i, j, 38)=RS("jiaA")
				tmpRec(i, j, 39)=RS("jiaB")
				tmpRec(i, j, 40)=RS("KZM")
				tmpRec(i, j, 41)=RS("jiaAM")
				tmpRec(i, j, 42)=RS("jiaBM")
				tmpRec(i, j, 43)=RS("BZKM")
				tmpRec(i, j, 44)=RS("KTAXM")
				tmpRec(i, j, 45)=RS("QITA")
				tmpRec(i, j, 46)=RS("FL")
				tmpRec(i, j, 47)=RS("real_total")
				tmpRec(i, j, 48)=RS("laonh")
				tmpRec(i, j, 49)=RS("sole")
				tmpRec(i, j, 50)=RS("dm")
				tmpRec(i, j, 51)=RS("lzbzj")
				tmpRec(i, j, 52)=RS("yymm")
				tmpRec(i, j, 53)=RS("outdate")
				tmpRec(i, j, 54)=RS("bh")
				tmpRec(i, j, 55)=cdbl(RS("hs")) 
				tmpRec(i, j, 56)=RS("GT")
				tmpRec(i, j, 57)=0
				tmpRec(i, j, 58)=cdbl(RS("BB"))+cdbl(RS("cv"))+cdbl(RS("phu"))+cdbl(RS("nn"))+cdbl(RS("kt"))+cdbl(RS("mt"))+cdbl(RS("ttkh"))
				tmpRec(i, j, 59)=cdbl(RS("wpamt"))
				tmpRec(i, j, 60)=cdbl(RS("wpamt"))+cdbl(RS("laonh"))
				tmpRec(i, j, 61)=cdbl(RS("H1M"))+cdbl(RS("H2M"))+cdbl(RS("H3M"))+cdbl(RS("b3M"))
				tmpRec(i, j, 62)=rs("dkm")
				tmpRec(i, j, 63)=cdbl(RS("jiaAM"))+cdbl(RS("jiaBM"))+cdbl(RS("KZM"))+cdbl(RS("BZKM"))				
				tmpRec(i, j, 64)=rs("memo")
				IF tmpRec(i, j, 64)<>"" AND isnull(rs("memo"))=false then 
					tmpRec(i, j, 64) = replace(tmpRec(i, j, 64),vbcrlf,"<br>")
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
	Session("empfileedit") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empfileedit")

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
 
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	'<%=self%>.empid1.focus()
	'<%=self%>.empid1.select()
end function

function datachg()
	<%=self%>.action="<%=self%>.foregnd.asp?totalpage=0"
	<%=self%>.submit
end function

 
</SCRIPT> 
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">


<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD >
	<img border="0" src="../image/icon.gif" align="absmiddle">
	薪資查詢
	</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<TABLE WIDTH=460 CLASS=FONT9 BORDER=0>
	<tr>
		<td  nowrap align=right>統計年月</td>
		<td nowrap colspan=3>
			<INPUT  NAME=yymm VALUE="<%=ym1%>" class=inputbox size=7>~
			<INPUT  NAME=yymm2 VALUE="<%=ym2%>" class=inputbox size=7>		
		</td>  
	</tr>
	<TR height=25 >
		<TD nowrap align=right >工號</TD>
		<TD ><%=empid1%> &nbsp; <%=empname%>			
		</TD>		
		<TD nowrap align=right >到職日</TD>
		<TD >
				<%=nindat%>
		</TD>				
	</TR>
	<TR height=25 >
		<TD nowrap align=right >稅號<br>MST</TD>
		<TD colspan=3><%=taxcode%>  	
		</TD>		
		 			
	</TR>	
	<!--TR>
		< TD nowrap align=right>簽約</TD>
		<TD >
			<select name=outemp class=font9 onchange="datachg()"> 
			 	<option value="" <%if outemp="" then %>selected<%end if%>>全部</option>
			 	<option value="Y" <%if outemp="Y" then %>selected<%end if%>>已簽約</option>
			 	<option value="N" <%if outemp="N" then %>selected<%end if%>>未簽約</option>
			 </select>	
		</TD> 
	</TR-->
</TABLE>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<!-------------------------------------------------------------------->
<TABLE CLASS="txt8" BORDER=0   cellspacing="1" cellpadding="1" bgcolor=#e4e4e4 >
 	<TR BGCOLOR="LightGrey" HEIGHT=25   >
 		<TD width=50 nowrap align=center>YYMM</TD>
 		<!--TD width=40  nowrap align=center>Quoc<BR>Tich</TD-->
 		<TD width=30 nowrap align=center>廠</TD>
 		<TD width=60 nowrap align=center>bo phan</TD> 		
 		<TD width=85 nowrap align=center>Chuc vu</TD>
 		<TD width=30 nowrap align=center>幣別</TD>
		<%if country="VN" then %>
		<TD width=60 nowrap align=center>基本<br>薪資</TD>
		<TD width=60 nowrap align=center>全勤</TD>
		<TD width=60 nowrap align=center>績效</TD>
		<TD width=60 nowrap align=center>其他<br>收入</TD>
		<TD width=60 nowrap align=center>加班費</TD>
		<TD width=60 nowrap align=center> Will<br>power</TD>
		<TD width=60 nowrap align=center>暫扣款</TD>
 		<TD width=60 nowrap align=center>(-)保險</TD>
 		<TD width=60 nowrap align=center>(-)工團</TD> 		
 		<TD width=60 nowrap align=center>(-)時假</TD>		
 		<TD width=60 nowrap align=center>(-)其他</TD> 		
 		<TD width=60 nowrap align=center>(-)所得稅</TD>
		<TD width=60 nowrap align=center>實領薪資</TD>
		<%else%>
		<TD width=40 nowrap align=center>基本<br>薪資</TD>
		<TD width=40 nowrap align=center>全勤</TD>
		<TD width=40 nowrap align=center>績效</TD>
		<TD width=40 nowrap align=center>其他<br>收入</TD>
		<TD width=40 nowrap align=center>加班費</TD>
		<TD width=40 nowrap align=center> Will<br>power</TD>
		<TD width=40 nowrap align=center>暫扣款</TD>
 		<TD width=40 nowrap align=center>(-)保險</TD>
 		<TD width=40 nowrap align=center>(-)工團</TD> 		
 		<TD width=40 nowrap align=center>(-)時假</TD>		
 		<TD width=40 nowrap align=center>(-)其他</TD> 		
 		<TD width=40 nowrap align=center>(-)所得稅</TD>
		<TD width=40 nowrap align=center>實領薪資</TD>
		<%end if%>
		<TD width=100 nowrap align=center>備註</TD>
		
 		 
 	</TR>
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN
			WKCOLOR="#ffffff"  '"LavenderBlush"
		ELSE
			WKCOLOR="#ffffff"  '"LavenderBlush"
		END IF
		if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
		<%if CurrentRow>1 and tmpRec(CurrentPage, CurrentRow-1, 1)<>tmpRec(CurrentPage, CurrentRow, 1) then %>
		<Tr>
			<Td bgcolor=black colspan=24></td>
		</tr>
		<%end if%>	
	<TR BGCOLOR='<%=WKCOLOR%>' height=22>
 		<TD align=center  >
 				<%=tmpRec(CurrentPage, CurrentRow, 52)%> 			
 		</TD>
		<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" name="f_empid" >
		<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 52)%>" name="f_yymm" >
 		<TD align=center  > <!--廠別-->
 			<%=tmpRec(CurrentPage, CurrentRow, 7)%>
 		</TD>
 		<TD align=LEFT nowrap ><!--shift+部門-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then%>
 				<%=tmpRec(CurrentPage, CurrentRow, 8)%>-<%=left(tmpRec(CurrentPage, CurrentRow, 12),10)%>
 			<%end if%>	
 		</TD>  
 		<Td><%=left(tmpRec(CurrentPage, CurrentRow, 14),10)%></td><!--職務-->
		<Td align="center"><%=tmpRec(CurrentPage, CurrentRow, 50)%></td><!--dm-->
 		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 58),0)%></td><!--基本薪-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 25),0)%></td><!--QC-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 28),0)%></td><!--jx-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 26),0)%></td><!--其他收入-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 61),0)%></td><!--all加班費-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 59),0)%></td><!--wp-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 62),0)%></td><!--暫扣-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 54),0)%></td><!--保險-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 56),0)%></td><!--工團-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 63),0)%></td><!--時假-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 45),0)%></td><!--扣其他-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 44),0)%></td><!--扣TAX-->
		<Td align="right"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 47),0)%></td><!--實領-->
		<Td align="left"><%=tmpRec(CurrentPage, CurrentRow, 64)%></td><!--memo-->
 		
	</TR> 
	<%end if%>
	<%next%>
</TABLE>
<input type=hidden value="" name="f_yymm" >
<input type=hidden value="" name="f_empid" >
<br>
<TABLE border=0 width=500   >
	<tr> 
	<td align="center">	 
		<input type="button" name="send" value="(X)關閉CLose"   class=button onclick="window.close()">
	</td>
	</TR>
</TABLE>
</form>




</body>
</html>

<script language=vbscript>
function BACKMAIN()

	open "empfile.fore1.asp" , "_self"
end function

function oktest(index)
	f1=<%=self%>.f_empid(index).value
	f2=<%=self%>.f_yymm(index).value
	'alert f1 & f2 
	'tp=<%=self%>.totalpage.value
	'cp=<%=self%>.CurrentPage.value
	'rc=<%=self%>.RecordInDB.value
	wt = (window.screen.width )*0.6
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02
	
	open "<%=self%>.showsalary.asp?empid="&f1&"&yymm="&f2  , "_blank", "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"	
	
end function

</script>

