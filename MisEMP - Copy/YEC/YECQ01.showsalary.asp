<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "YECQ01S"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

ym = request("yymm")
empid1 = trim(REQUEST("empid"))

ccnt = cdbl(ym2) - cdbl(ym1)
if ccnt = 0 then ccnt = 1 

gTotalPage = 1
PageRec = 20*ccnt   'number of records per page
TableRec = 60    'number of fields per record
'NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/01"
NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)

 
	sql="select isnull(e.lj,'') lj , isnull(e.ljstr,'') ljstr ,  "&_
		"isnull(d.lw,'') lw , isnull(d.lg,'') lg , isnull(d.lz,'') lz , isnull(d.ls,'') ls , "&_
		"isnull(d.lwstr,'') lwstr , isnull(d.lgstr,'') lgstr , isnull(d.lzstr,'') lzstr, isnull(d.lsstr,'') lsstr, "&_
		"c.empnam_cn, c.empnam_vn, c.nindat, c.outdate, a.* , isnull(b.real_total,0) as backTotal , f.sys_value as Sjstr , isnull(g.totamt,0) wpamt,  "&_
		"isnull(g.bb,0) as wp_bb, isnull(g.cv,0) as wp_cv, isnull(g.rzm,0) rzm,  isnull(g.rzdays,0) rzdays, isnull(g.jrm,0) jrm, isnull(g.jrdays,0) jrdays,  "&_
		"isnull(g.tnkh,0) wp_tnkh, "&_
		"isnull(g.qita,0) wp_qita, isnull(g.zkm,0) wp_zkm, isnull(g.zgm,0) wp_zgm, isnull(g.memo,'') wp_memo , isnull(a.memo,'') salary_memo "&_
		"from "&_		
		"(select  empid, empnam_cn, empnam_vn, convert(char(10),indat,111)  nindat , isnull(convert(char(10),outdat,111),'')  outdate  from empfile  where empid = '"&empid1&"' ) c   "&_
		"left join ( select * from empdsalary where empid = '"&empid1&"' and yymm='"&ym&"' ) a on a.empid= c.empid  "&_
		"left join ( select * from empdsalary_bak   )  b on b.yymm = a.yymm and b.empid = c.empid  "&_		
		"left join (select *from view_empgroup  ) d on d.empid = c.empid  and d.yymm = a.yymm   "&_
		"left join (select *from view_empjob) e on e.empid = c.empid and e.yymm = a.yymm "&_
		"left join (select *from basicCode) f on f.sys_type = a.job "&_
		"left join (select * from salarywp ) g on g.empid = c.empid and g.yymm = a.yymm "  
sql = sql & "order by c.empid , a.yymm " 
'response.write sql 
	rs.Open sql, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
		country = rs("country")
		lw = rs("lw")
		groupid = rs("lg")
		
		empid= trim(rs("empid"))
		empnam_cn= trim(rs("empnam_cn"))
		empnam_vn = trim(rs("empnam_vn"))		
		nindat= rs("nindat")
		lj = rs("lj")
		lw = rs("lw")
		ls = rs("ls")
		lg = RS("lg")
		lz = RS("lz")
		lwstr=RS("lwstr")				
		lgstr=RS("lgstr")
		lzstr=RS("lzstr")
		ljstr=RS("ljstr")
		job=RS("job")
		whsno=RS("whsno")
		Sjstr=RS("Sjstr") 				
		bb=RS("BB")
		cv=RS("CV")
		phu=RS("PHU")
		nn=RS("NN")
		kt=RS("KT")
		mt=RS("MT")
		ttkh=RS("TTKH")
		qc=RS("QC")
		tnkh=RS("TNKH")
		tbtr=RS("TBTR")
		jx=RS("JX")
		h1m=RS("H1M")
		h2m=RS("H2M")
		h3m=RS("H3M")
		b3m=RS("B3M")
		h1=RS("H1")
		h2=RS("H2")
		h3=RS("H3")
		b3=RS("B3")
		kzhour=RS("kzhour")
		jiaA=RS("jiaA")
		jiaB=RS("jiaB")
		KZM=RS("KZM")
		jiaAM=RS("jiaAM")
		jiaBM=RS("jiaBM")
		BZKM=RS("BZKM")
		KTAXM=RS("KTAXM")
		QITA=RS("QITA")
		FL=RS("FL")
		real_total=RS("real_total")
		laonh=RS("laonh")
		sole=RS("sole")
		dm=RS("dm")
		lzbzj=RS("lzbzj")
		yymm=RS("yymm")
		outdate=RS("outdate")
		bh=RS("bh")
		hs=cdbl(RS("hs")) 
		gt=RS("GT")
		
		all_bsaic_salary=cdbl(RS("BB"))+cdbl(RS("cv"))+cdbl(RS("phu"))+cdbl(RS("nn"))+cdbl(RS("kt"))+cdbl(RS("mt"))+cdbl(RS("ttkh"))+cdbl(RS("qc"))
		tot_salary_jia = cdbl(all_bsaic_salary) + cdbl(tnkh)+cdbl(tbtr)+cdbl(h1m)+cdbl(h2m)+cdbl(h3m)++cdbl(b3m)
		totamt=cdbl(RS("wpamt"))+cdbl(RS("laonh"))		  '境內+境外		
		wp_amt=cdbl(RS("wpamt"))
		wp_bb=cdbl(RS("wpamt"))
		wp_cv=cdbl(RS("wpamt"))
		rzm=(RS("rzm"))
		rzdays=(RS("rzdays"))
		jrm=(RS("jrm"))
		jrdays=(RS("jrdays"))
		wp_tnkh = rs("wp_tnkh")
		wp_qita = rs("wp_qita")
		wp_zkm = rs("wp_zkm")
		wp_zgm = rs("wp_zgm")
		wp_memo = rs("wp_memo")
		salary_memo = rs("salary_memo")
		acc = rs("acc")
		
		money_h =rs("money_h") 
		real_total = rs("real_total")
		
		tot_salary_jan = cdbl(KZM)+cdbl(jiaaM)+cdbl(jiaBM)+cdbl(BZKM)+cdbl(ktaxm)+cdbl(QITA)+cdbl(bh)+cdbl(hs)+cdbl(GT) 
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
<body  topmargin="10" leftmargin="10"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
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
	<%=session("pgname")%> 
	</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<TABLE WIDTH=460 CLASS=FONT9 BORDER=0>
	<tr>
		<td  nowrap align=right>統計年月</td>
		<td nowrap colspan=3>
			<INPUT  NAME=yymm VALUE="<%=ym%>" class="readonly" size=10 readonly >
		</td>
		<TD nowrap align=right>廠別</TD>
		<TD >
			<select name=WHSNO  class=txt8 onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=lw THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>		
		<TD nowrap align=right >國籍</TD>
		<TD >
			<select name=COUNTRY  class=txt8  onchange="datachg()" >
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>			
		</TD>
		<TD nowrap align=right >部門</TD>
		<TD >
			<select name=GROUPID  class=txt8  onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=GROUPID THEN %> SELECTED <%END IF%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
	</TR> 
</TABLE>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<!-------------------------------------------------------------------->
<TABLE CLASS="txt8" BORDER=0   cellspacing="1" cellpadding="1" bgcolor=#e4e4e4 >
 	<TR BGCOLOR="LightGrey" HEIGHT=25   >
 		<TD width=30 nowrap align=center>廠</TD>
 		<TD width=60 nowrap align=center>bo phan</TD>
 		<TD width=45 nowrap align=center>工號<br>Ma So</TD>
 		<TD width=120 nowrap align=center>姓名<br>Ho Ten</TD>
 		<TD width=70 nowrap align=center>到職日期<br>NVX</TD>
 		<TD width=70 nowrap align=center>Chuc vu</TD> 		
	</tr> 		
 	</TR>
	<TR BGCOLOR="#ffffff" height=22> 				 		
 		<TD align=center  > <!--廠別-->
 			<%=whsno%>
 		</TD>
 		<TD align=LEFT nowrap ><!--shift+部門--> 			
 			<%=ls%>-<%=left(lgstr,3)%> 			
 		</TD> 
 		<TD nowrap align=center><%=empid%></TD> 		
 		<TD nowrap>
 				<%=empnam_cn%><BR><font class=txt8VN><%=empnam_vn%></font> 			
 		</TD>
 		<TD align=center nowrap>
 			<%=nindat%><BR><font color=red><%=outdate%></font>
 		</TD>
 		<Td><%=left(ljstr,6)%></td> 		
 		</td>		
	</TR> 
</TABLE><br>
<TABLE CLASS="txt8" BORDER=0   cellspacing="1" cellpadding="1" bgcolor=#e4e4e4 >
 	<TR BGCOLOR="LightGrey" HEIGHT=25   > 		
		<TD width=30 nowrap align=center>立帳<br>單位</TD>
 		<TD width=30 nowrap align=center>幣別</TD> 		
		<TD width=60 nowrap align=center>(+)BB</TD>
		<TD width=60 nowrap align=center>(+)CV</TD>
 		<TD width=60 nowrap align=center>(+)PHU</TD>
		<TD width=60 nowrap align=center>時薪</TD>
 		<TD width=60 nowrap align=center>(+)NN</TD> 		
		<TD width=60 nowrap align=center>(+)KT</TD> 		
		<TD width=60 nowrap align=center>(+)MT</TD> 		
		<TD width=60 nowrap align=center>(+)TTKH</TD> 		
		<TD width=60 nowrap align=center>(+)QC</TD>		
		<TD width=60 nowrap align=center>SALARY</TD>		
		<TD width=60 nowrap align=center>(+)TNKH</TD>
		<TD width=60 nowrap align=center>(+)TBTR</TD>		
	</tr> 	
	<TR BGCOLOR="#ffffff" height=25> 				 		
		<TD align=center  ><%=acc%>&nbsp;</TD>
		<TD align=center  ><%=dm%></TD>
 		<TD align=right  ><%=formatnumber(bb,0)%></TD>
		<TD align=right  ><%=formatnumber(cv,0)%></TD>
		<TD align=right  ><%=formatnumber(phu,0)%></TD>
		<TD align=right  bgcolor="lightyellow" ><%=formatnumber(money_h,0)%></TD>
		<TD align=right  ><%=formatnumber(nn,0)%></TD>
		<TD align=right  ><%=formatnumber(kt,0)%></TD>
		<TD align=right  ><%=formatnumber(mt,0)%></TD>
		<TD align=right  ><%=formatnumber(ttkh,0)%></TD>		
		<TD align=right  ><%=formatnumber(qc,0)%></TD>
		<TD align=right  bgcolor="lightyellow"><%=formatnumber(all_bsaic_salary,0)%></TD>
		<TD align=right  ><%=formatnumber(TNKH,0)%></TD>
		<TD align=right  ><%=formatnumber(TBTR,0)%></TD>		
	</tr>
	<TR BGCOLOR="LightGrey" HEIGHT=25   >
		<TD width=60 nowrap align=center  colspan=2>&nbsp;</td>
		<TD width=60 nowrap align=center>(+)H1M</TD>
		<TD width=60 nowrap align=center>(+)H2M</TD>
		<TD width=60 nowrap align=center>(+)H3M</TD>
		<TD width=60 nowrap align=center>(+)B3M</TD>	
		<TD width=60 nowrap align=center>(+)TOT</TD>	
		<TD width=60 nowrap align=center>&nbsp;</TD>	
		<TD width=60 nowrap align=center>&nbsp;</TD>	
		<TD width=60 nowrap align=center>&nbsp;</TD>	
		<TD width=60 nowrap align=center>&nbsp;</TD>	
		<TD width=60 nowrap align=center>實領金額</TD>	
		<TD width=60 nowrap align=center>整數</TD>	
		<TD width=60 nowrap align=center>零數</TD>		
	</tr>
	<TR BGCOLOR="white" HEIGHT=25   >
		<TD width=60 nowrap align=center  colspan=2>&nbsp;</td>
		<TD align=right  ><%=formatnumber(H1M,0)%><br>(<%=h1%>)</TD>
		<TD align=right  ><%=formatnumber(H2M,0)%><br>(<%=h2%>)</TD>
		<TD align=right  ><%=formatnumber(H3M,0)%><br>(<%=h3%>)</TD>
		<TD align=right  ><%=formatnumber(B3M,0)%><br>(<%=b3%>)</TD>
		<TD align=right  ><font color="red"><%=formatnumber(tot_salary_jia,0)%></font></TD>
		<TD align=right  >&nbsp;</TD>
		<TD align=right  >&nbsp;</TD>
		<TD align=right  >&nbsp;</TD>
		<TD align=right  >&nbsp;</TD>
		<TD align=right  ><font color="red"><%=formatnumber(real_total,0)%></font></TD>
		<TD align=right  ><font color="red"><%=formatnumber(laonh,0)%></font></TD>
		<TD align=right  ><font color="red"><%=formatnumber(sole,0)%></font></TD>
	</tr>
 	<TR BGCOLOR="LightGrey" HEIGHT=25   >
 		<TD width=60 nowrap align=center  colspan=2>&nbsp;</td>
		<TD width=60 nowrap align=center >(-)保險</TD>
		<TD width=60 nowrap align=center>(-)工團</TD>
		<TD width=60 nowrap align=center>(-)伙食</TD>		
		<TD width=60 nowrap align=center>(-)不足月</TD>
 		<TD width=60 nowrap align=center>(-)所得稅</TD>  
		<TD width=60 nowrap align=center>(-)住宿</TD>
		<TD width=60 nowrap align=center>(-)曠職</TD>
		<TD width=60 nowrap align=center>(-)事假</TD> 		
		<TD width=60 nowrap align=center>(-)病假</TD> 		
		<TD width=60 nowrap align=center>(-)TOT</TD>
 		
	</tr> 	
	<TR BGCOLOR="#ffffff" height=25> 
		<TD width=60 nowrap align=center  colspan=2>&nbsp;</td>
		<TD align=right  ><%=formatnumber(bh,0)%></TD>
		<TD align=right  ><%=formatnumber(gt,0)%></TD>
 		<TD align=right  ><%=formatnumber(hs,0)%></TD>
		<TD align=right  ><%=formatnumber(bzkm,0)%></TD>
		<TD align=right  ><%=formatnumber(ktaxm,0)%></TD>	 
		<TD align=right  >&nbsp;</TD>
		<TD align=right  ><%=formatnumber(kzm,0)%><br>(<%=kzhour%>)</TD>
		<TD align=right  ><%=formatnumber(jiaaM,0)%><br>(<%=jiaA%>)</TD>
		<TD align=right  ><%=formatnumber(jiabM,0)%><br>(<%=jiaB%>)</TD>
		<TD align=right  ><font color="red"><%=formatnumber(tot_salary_jan,0)%></font></TD>	 
	</tr>	
	<tr height=25 bgcolor="#ffffff">
		<td colspan=2 align="right">memo:</td>
		<td colspan=12><%=salary_memo%></td>
	</tr>
</TABLE>	<br>
<TABLE CLASS="txt8" BORDER=0   cellspacing="1" cellpadding="1" bgcolor=#e4e4e4 >
 	<TR BGCOLOR="LightGrey" HEIGHT=25   >
 		<TD width=60 nowrap align=center >境外</td>
		<TD width=60 nowrap align=center >(+)職加</TD>
		<TD width=60 nowrap align=center>(+)海外<br>津貼</TD>
		<TD width=60 nowrap align=center>(+)非假日<br>(差旅)</TD>		
		<TD width=60 nowrap align=center>(+)假日<br>(差旅)</TD>
 		<TD width=60 nowrap align=center>(+)其他收入</TD>  
		<TD width=60 nowrap align=center>(-)其他</TD>
		<TD width=60 nowrap align=center>(-)暫扣款</TD>		
		<TD width=60 nowrap align=center>TOT</TD>
		<TD width=60 nowrap align=center >&nbsp;</td>
		<TD width=60 nowrap align=center >境內+境外</td>
 		
	</tr> 	
	<TR BGCOLOR="#ffffff" height=25> 
		<TD width=60 nowrap align=center >&nbsp;</td>
		<TD align=right  ><%=formatnumber(wp_bb,0)%></TD>
		<TD align=right  ><%=formatnumber(wp_cv,0)%></TD>
 		<TD align=right  ><%=formatnumber(rzm,0)%><br>(<%=rzdays%>)</TD>
		<TD align=right  ><%=formatnumber(jrm,0)%><br>(<%=jrdays%>)</TD>		
		<TD align=right  ><%=formatnumber(wp_tnkh,0)%></TD>	 
		<TD align=right  ><%=formatnumber(wp_qita,0)%></TD>	 		
		<TD align=right  ><%=formatnumber(wp_zkm,0)%></TD>	 		
		<TD align=right  ><font color="red"><%=formatnumber(wp_amt,0)%></font></TD>	 
		<TD width=60 nowrap align=center >&nbsp;</td>
		<TD align=right  ><font color=red><%=formatnumber(totamt,0)%></font></TD>	 		
	</tr>	
	<tr height=22 bgcolor="#ffffff">
		<td align="right">memo:</td>
		<Td colspan=10><%=wp_memo%></td>
	</tr>
</TABLE>	
 
 <br>

<TABLE border=0 width=500 class=font9 >
	<tr>  
	<td align="center">	
		<input type="button" name="send" value="關閉視窗Close"   class=button onclick="window.close()">
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

