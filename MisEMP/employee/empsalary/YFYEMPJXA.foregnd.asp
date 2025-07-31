<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<%
self="YFYEMPJXA"

Set conn = GetSQLServerConnection()

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth = year(date())&right("00"&month(date())-1,2)
end if

if day(date())<=11 then
	if month(date())="01" then
		calcmonth = year(date()-1)&"12"
	else
		calcmonth = year(date())&right("00"&month(date())-1,2)
	end if
else
	calcmonth = nowmonth
end if


JXYM=REQUEST("JXYM")
salaryYM=REQUEST("YYMM")
country=REQUEST("country")
JOBID=REQUEST("JOBID")
empid1 = REQUEST("empid1")
GROUPID = REQUEST("GROUPID")
SHIFTN=REQUEST("SHIFT")
zuno = REQUEST("zuno")

'response.write groupid &"<BR>"
JXYMdays=left(jxym,4)&"/"&right(jxym,2)&"/01"
Edays=left(salaryYM,4)&"/"&right(salaryYM,2)&"/01" 

e_edat = left(salaryYM,4)&"/"&right(salaryYM,2)&"/25"

SQLSTR ="SELECT B.SYS_VALUE AS GSTR , C.SYS_VALUE AS ZSTR , A.* FROM " &_
		"(SELECT * FROM YFYMJIXO where JXYM='"& JXYM &"' AND GROUPID='"& GROUPID &"' and isnull(zuno,'')  like '"& zuno &"%' AND SHIFT like '%"& SHIFTN &"' ) A "&_
		"LEFT JOIN (select * from basicCode where func ='groupid' ) B ON B.SYS_TYPE = A.GROUPID "&_
		"LEFT JOIN (select * from basicCode where func ='zuno' ) C ON C.SYS_TYPE = isnull(A.zuno,'') "&_
		"ORDER BY A.STT "
Set rDs = Server.CreateObject("ADODB.Recordset")
RDS.OPEN SQLSTR,CONN, 3, 3
'RESPONSE.WRITE "1="&SQLSTR &"<br>"

IF NOT RDS.EOF THEN
	CurrentPage = 1
	ICOUNT = RDS.RecordCount
	GSTR = RDS("GSTR")
	ZSTR = RDS("ZSTR")
	Redim ARRAYS(ICOUNT, 11)  'Array
	for i = 1 to ICOUNT
		ARRAYS(i,1)=rDs("JXYM")
		ARRAYS(i,2)=rDs("salaryYM")
		ARRAYS(i,3)=rDs("groupID")
		ARRAYS(i,4)=rDs("shift")
		ARRAYS(i,5)=rDs("STT")
		ARRAYS(i,6)=TRIM(rDs("DESCP"))
		ARRAYS(i,7)=rDs("HXSL")
		ARRAYS(i,8)=rDs("HESO")
		ARRAYS(i,9)=rDs("PER")
		ARRAYS(i,10)=rDs("AUTOID")
		ARRAYS(i,11)=rDs("zuno")
		RDS.MOVENEXT
	next
	Session("EMPJX") = ARRAYS
else
%>	<SCRIPT LANGUAGE=VBS>
		ALERT "無當月績效資料!!"
		OPEN "<%=SELF%>.ASP", "Fore"
	</SCRIPT>
<%
	response.end 
END IF

TotalPage = 10
PageRec = 10  'number of records per page
TableRec = 30  'number of fields per record

salaryYMDate = left(salaryYM,4)&"/"&right(salaryYM,2)&"/01"
days = DAY(cDatestr+(32-DAY(salaryYMDate))-DAY(cDatestr+(32-DAY(salaryYMDate))))   '月有幾天
enddays=left(salaryYM,4)&"/"&right(salaryYM,2)&"/"&right("00"&cstr(days),2)

JXYMdays=left(JXYM,4)&"/"&right(JXYM,2)&"/01"
JXdays = DAY(cDatestr+(32-DAY(JXYMdays))-DAY(cDatestr+(32-DAY(JXYMdays))))   '月有幾天
JXenddays=left(JXYM,4)&"/"&right(JXYM,2)&"/"&right("00"&cstr(JXdays),2)

'response.write "JXenddays=" & JXenddays &"<BR>"
'response.write "Edays=" & Edays &"<BR>" 

SQL=""
SQL=SQL&"select isnull(kh.monfen,0) monfen , z.oldjob, z1.sys_value as oldjobdesc, isnull(z2.bonus,0)  as Jxbonus, y.LG as OldG, "
SQL=SQL&"Y.LS, Y.LZ , a.bhdat, isnull(b.SUKM,0) SUmoney , isnull ( ( c.latefor+c.forget ) , 0 ) as NFL, "
sql=sql&"c.latefor, c.forget, isnull(c.KZHOUR,0) kzhour , a.* from "
SQL=SQL&"( select * from view_empfile where country<>'TW' and country<>'MA'  ) a  "
SQL=SQL&"left join (  "
SQL=SQL&"SELECT  CFGROUP,  EMPID, CFDW,  SUM(SUKM*EXRT) SUKM FROM (  "
SQL=SQL&"SELECT isnull(B.EXRT,1) exrt,   A.* FROM  "
SQL=SQL&"( SELECT * FROM  yfydsuco WHERE  YM='"& salaryYM &"'  ) A  "
SQL=SQL&"LEFT JOIN ( SELECT *  FROM VYFYEXRT  WHERE  YYYYMM='"& salaryYM &"' ) B ON B.CODE = A.DM  "
SQL=SQL&") Z GROUP BY  CFGROUP,  EMPID, CFDW  HAVING EMPID<>'' "
SQL=SQL&") b on b.empid = a.empid "
SQL=SQL&"left join ( "
SQL=SQL&"select EMPID ,  sum( KZHOUR) as KZHOUR, SUM(latefor) latefor, SUM(FORGET) FORGET from empwork where yymm='"& JXYM &"' GROUP BY EMPID "
SQL=SQL&") c on c.empid = a.empid "
SQL=SQL&"LEFT JOIN( SELECT * FROM View_empGroup where LW='LA' and yymm='"& JXYM &"'   ) y ON  y.EMPID= A.EMPID "  
SQL=SQL&"LEFT JOIN(  select  yymm, empid , max(job) as oldjob from bempj  where   yymm='"& JXYM &"' group by  yymm, empid  ) z ON z.YYMM='"& JXYM &"' AND z.EMPID= A.EMPID     "
SQL=SQL&"LEFT JOIN( SELECT * FROM basicCode where  func='LEV' ) z1  ON  z1.sys_type = z.oldjob    " 
SQL=SQL&"left join ( select *  from  empsalarybasic where func='DD'  )z2 on z2.job = z.oldjob and z2.country = a.country      " 
SQL=SQL&"left join (select khym, empid, sum(fnA+fnb+fnc+fnd) as monfen  from empkhb where  khym='"&JXYM&"' group by khym, empid ) kh on kh.empid = a.empid  and kh.khym = '"&JXYM&"' "
SQL=SQL&"where isnull(a.nindat,'')<='"& JXYMdays &"' and  (isnull(outdat,'')='' or isnull(outdat,'')<>'' and rtrim(ltrim(outdate))>'"& JXenddays &"' )   "
sql=SQL&"AND y.LG='"&GROUPID &"' and  isnull(Y.LS,'') like'%"& SHIFTN &"'  "
sql=sql&"and lz<= case when lg='A065' then 'A0655'  else 'zzzzz' end AND y.LZ LIKE '"& zuno &"%'  " 
SQL=SQL&"order by a.empid "  
'RESPONSE.WRITE SQL &"<BR>"
'RESPONSE.END
Set rs = Server.CreateObject("ADODB.Recordset")
RS.OPEN SQL,CONN, 3, 3

IF NOT RS.EOF THEN
pagerec = rs.RecordCount
rs.PageSize = pagerec
RecordInDB = rs.RecordCount
TotalPage = rs.PageCount
Redim tmpRec(TotalPage,PageRec, TableRec)  'Array
for i = 1 to TotalPage
	for j = 1 to PageRec
		if not rs.EOF then
			tmpRec(i, j, 0)="no"
			tmpRec(i, j, 1)=trim(rs("EMPID"))
			tmpRec(i, j, 2)=trim(rs("OldG"))  '績效年月時之單位
			tmpRec(i, j, 3)=trim(rs("LS"))  '績效年月時之班別
			tmpRec(i, j, 4)=rs("Jxbonus")
			tmpRec(i, j, 5)=rs("EMPNAM_CN")
			tmpRec(i, j, 6)=rs("EMPNAM_VN")
			tmpRec(i, j, 7)=rs("nindat")
			if JXYM>"200702" then 
				if  rs("country")="CN" then 
					tmpRec(i, j, 8) = 0
				else
					tmpRec(i, j, 8)= cdbl(RS("KZHOUR"))
				end if 	
			else
				if  rs("country")="CN" then 
					tmpRec(i, j, 8) = 0
				else
					tmpRec(i, j, 8)= rs("latefor")+rs("forget")
				end if 	
			end if	
			tmpRec(i, j, 9)=rs("SUmoney")
			tmpRec(i, j, 10)= rs("oldjob")
			tmpRec(i, j, 11)=left(rs("oldjobdesc"),4)

			F1_TOTJX = 0
			for z = 1 to icount
				'共有幾項績效項目
				tmpRec(i, j,11+z)= round( (cdbl(rs("Jxbonus"))* ( cdbl(arrays(z,9))/100 ) ) * cdbl(arrays(z,8)) , 0) 
				F1_TOTJX = F1_TOTJX + cdbl(tmpRec(i, j, 11+z))
				'response.write tmpRec(i, j, 1) &"-" & 11+Z &"-" & tmpRec(i, j, 11+z) &"-" & arrays(z,9) &"-" &arrays(z,8) &"<BR>"
			next
			
			if tmpRec(i, j, 8)< 8 then
				tmpRec(i, j, 12+icount)= 0
			elseif cdbl(tmpRec(i, j, 8))>=8 and cdbl(tmpRec(i, j, 8))<16 then
				tmpRec(i, j,12+icount) = round(cdbl(F1_TOTJX)*0.5,0)'曠職ㄧ天以上...扣績效獎金一半
			elseif cdbl(tmpRec(i, j, 8))>=16 then
				tmpRec(i, j, 12+icount)= cdbl(F1_TOTJX)  ''曠職2天以上...無績效獎金
			end if
			
			tmpRec(i, j, 13+icount) = F1_TOTJX
			'tmpRec(i, j, 14+icount) = round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,12+icount))-cdbl(tmpRec(i, j, 9)),0)
			tmpRec(i, j, 15+icount) = rs("outdate")
			IF RS("COUNTRY")="CN" THEN 
				tmpRec(i, j, 16+icount) = 1 
			ELSE
				'if cdbl(rs("EMPPER"))>=1 then 
				'	tmpRec(i, j, 16+icount) = 1  'rs("EMPPER") 
				'	'tmpRec(i, j, 16+icount) =1 ' rs("EMPPER") 
				'else
					tmpRec(i, j, 16+icount) = 1  'rs("EMPPER")
				'end if 	
			END IF 	
			if cdbl(tmpRec(i, j, 16+icount))>=1 then 
				tmpRec(i, j, 17+icount) =  fix( round( tmpRec(i, j, 14+icount),0) /1000) * 1000 
			else	
				tmpRec(i, j, 17+icount)  =  fix( round( cdbl(tmpRec(i, j, 16+icount))* cdbl(tmpRec(i, j,14+icount)) ,0) /1000) * 1000 
			end if 	
			
			IF left(replace(rs("outdate"),"/",""),6)>=JXYM and trim(rs("outdate"))< e_edat  then 
				tmpRec(i, j, 14+icount) = 0
			else
				if jxym>"200702" then 
					if cdbl(tmpRec(i, j, 16+icount))>=1 then 
						tmpRec(i, j, 14+icount) =  ( round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,12+icount))-cdbl(tmpRec(i, j, 9)),0)) 
					else	
						tmpRec(i, j, 14+icount)  =  round(  cdbl(tmpRec(i, j, 16+icount)) *round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,12+icount))-cdbl(tmpRec(i, j, 9)),0) ,0)
					end if 	
				else
					tmpRec(i, j, 14+icount) = round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,12+icount))-cdbl(tmpRec(i, j, 9)),0) 
				end if 
			end if 			
			tmpRec(i, j, 18+icount ) = rs("groupid")  '目前單位
			if rs("shift")="" then 
				tmpRec(i, j, 19+icount ) = rs("LS")
			else
				tmpRec(i, j, 19+icount ) = rs("shift")  '目前班別
			end if	
			
			tmpRec(i, j, 20+icount ) = rs("zuno") '目前組別
			tmpRec(i, j, 21+icount ) =rs("LZ")  '績效年月的組別
			tmpRec(i, j, 22+icount ) =rs("monfen")  '考核分數			
			rs.MoveNext
		else
			exit for
		end if
	next
NEXT
Session("YFYEMPJXM") = tmpRec

ELSE %>
	<SCRIPT LANGUAGE=VBS>
	ALERT "員工資料有誤!!"
	'OPEN "<%=SELF%>.ASP", "Fore"
	</SCRIPT>
<%	 RESPONSE.END
END IF
%>
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
</head>
<body  topmargin="5" leftmargin="5" marginwidth="0" marginheight="0" onkeydown="enterto()" >
<form name="<%=self%>" method="post" action="<%=self%>.upd.asp">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden  NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME="ICOUNT"   VALUE="<%=ICOUNT%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<TD><img border="0" src="../../image/icon.gif" align="absmiddle"> 績效獎金</TD>
	</tr>
</table>
<hr size=style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=500><TR><TD ALIGN=CENTER>
	<table  width=500  class=txt9>
		<tr>
			<td ALIGN=LEFT>績效年月</td>
			<td><INPUT NAME="JXYM" VALUE="<%=JXYM%>" CLASS=READONLY2 READONLY SIZE=8 ></td>
			<td ALIGN=RIGHT>計薪年月</td> <td><INPUT NAME="SALARYYM" VALUE="<%=SALARYYM%>"  CLASS=READONLY2  READONLY  SIZE=8   ></td>
			<td ALIGN=RIGHT>單位</td>
			<td>
				<INPUT type=hidden NAME="GID" VALUE="<%=GROUPID%>" CLASS=READONLY2 READONLY size=5 >
				<INPUT NAME="GSTR" VALUE="<%=GSTR%>" CLASS=READONLY2 READONLY SIZE=8 >
				<INPUT  type=hidden  NAME="ZID" VALUE="<%=Zuno%>" CLASS=READONLY2 READONLY size=5 >
				<INPUT NAME="ZSTR" VALUE="<%=ZSTR%>" CLASS=READONLY2 READONLY SIZE=8 >
			</td>
			<td ALIGN=RIGHT>班別</td>
			<td>
				<INPUT NAME="shiftn" VALUE="<%=SHIFTN%>" CLASS="READONLY2" READONLY SIZE=4 >
			</td>
		</tr>
		<TR>
		<TD COLSPAN=8 HEIGHT=5></TD>
		</TR>
	</table>
	<table width=500 class=txt9 BORDER=0 cellspacing="1" cellpadding="2" BGCOLOR="#CCCCCC" >
		<tr bgcolor="#FFFFCC">
			<TD HEIGHT=22 ALIGN=CENTER>項次<br>比例</TD>
			<%FOR II=1 TO ICOUNT %>
				<TD ALIGN=CENTER CLASS=TXT8 nowrap ><%=ARRAYS(II,5)%><%=ARRAYS(II,6)%><br><%=ARRAYS(II,9)%>%</TD>
			<%NEXT%>
		</tr>
		<tr BGCOLOR="#CEE7FF">
			<TD HEIGHT=22 ALIGN=CENTER >實績</TD>
			<%FOR II=1 TO  ICOUNT %>  <TD ALIGN=CENTER><%=FORMATNUMBER(ARRAYS(II,7),2)%></TD>
			<%NEXT%>
		</tr>
		<tr BGCOLOR="#FED9CF">
			<TD HEIGHT=22 ALIGN=CENTER>係數</TD>
			<%FOR II=1 TO ICOUNT %> <TD ALIGN=CENTER><%=ARRAYS(II,8)%></TD> <%NEXT%>
		</tr>
	</table>
</TD></TR></TABLE>
<hr size=0	 style='border: 1px dotted #999999;'align=left width=500>
<table width=650 border=0 ><tr><td>
	<TABLE width=600 CLASS=TXT8 BGCOLOR="#CCCCCC" BORDER=0 border="1" cellspacing="1" ALIGN=CENTER>
		<TR BGCOLOR="#FEF7CF">
			<TD width=50 HEIGHT=22 ALIGN=CENTER nowrap>工號</TD>
			<TD width=120 ALIGN=CENTER nowrap>姓名</TD>
			<TD width=70 ALIGN=CENTER nowrap>到職日</TD>
			<TD width=80 ALIGN=CENTER  nowrap>職等</TD>
			<TD width=40 ALIGN=CENTER nowrap>單位</TD>
			<%FOR J=1 TO ICOUNT%>
				<TD ALIGN=CENTER width=60 CLASS=TXT8 nowrap><%=ARRAYS(J,5)%><br><%=ARRAYS(J,6)%></TD>
			<%NEXT%>
			 <td  width=100 align=center>小計</td>
			 <td width=60 align=center>曠職<BR>時數</td>
			 <td width=60 align=center>扣款金額　</td>
			 <td width=50 nowrap align=center>反規定</td>
			 <td width=60 nowrap align=center>事故扣款</td>
			 <td width=50 nowrap align=center>考核</td>
			 <td width=70 nowrap align=center>合計</td>
			 
		</TR>
			<%for CurrentRow = 1 to PageRec
				IF CurrentRow MOD 2 = 0 THEN
					WKCOLOR="#FFFFFF"
				ELSE
					WKCOLOR="#FFFFFF"
				END IF
				if tmpRec(CurrentPage,  CurrentRow,  1)<>"" then
			%>
					<TR id=id1 bgColor="<%=wkcolor%>" >
						<TD nowrap align=center style="cursor:hand" onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)"  >
							<%=tmpRec(CurrentPage,CurrentRow, 1)%>
							<input type=hidden name="empid" value="<%=tmpRec(CurrentPage,CurrentRow, 1)%>">
							<input type=hidden name="NowGroup" value="<%=tmpRec(CurrentPage,CurrentRow, 18+icount)%>">
							<input type=hidden name="NowShift" value="<%=tmpRec(CurrentPage,CurrentRow, 19+icount)%>">
							<input type=hidden name="NowZuno" value="<%=tmpRec(CurrentPage,CurrentRow, 20+icount)%>">
							<input type=hidden name="jxgroup" value="<%=tmpRec(CurrentPage,CurrentRow, 2)%>">
							<input type=hidden name="jxShift" value="<%=tmpRec(CurrentPage,CurrentRow, 3)%>">
							<input type=hidden name="jxZuno" value="<%=tmpRec(CurrentPage,CurrentRow, 21+icount)%>">
						</TD>
						<TD nowrap style="cursor:hand" onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)"><%=tmpRec(CurrentPage,CurrentRow, 5)%><BR><%=tmpRec(CurrentPage, CurrentRow, 6)%></TD>
						<TD nowrap align=center ><%=tmpRec(CurrentPage,CurrentRow,7)%><br><font color=red><%=tmpRec(CurrentPage,CurrentRow,15+icount)%></font></TD>
						<TD nowrap><%=tmpRec(CurrentPage,CurrentRow,11)%></TD>
						<TD nowrap align=center><%=tmpRec(CurrentPage, CurrentRow, 4)%></TD>
						<%for xx = 1 to ICOUNT %>
							<TD nowrap align=center> <%'response.write CurrentRow &"-"&cdbl(11)+cdbl(xx)  %>
								<input name="JX<%=arrays(xx,5)%>"  class="inputbox8r" size="6" value="<%=tmpRec(CurrentPage,CurrentRow,cdbl(11)+cdbl(xx))%>"  onchange="datachg(<%=currentRow-1%>)" >
								<input type=hidden size=3 name="BJX<%=arrays(xx,5)%>" value="<%=tmpRec(CurrentPage, CurrentRow,cdbl(11)+cdbl(xx))%>" >
							</TD>
						<%next%>
						<TD nowrap align=center><input name="sumjx" class="readonly8s" size=7 value="<%=tmpRec(CurrentPage, CurrentRow,13+icount)%>" readonly style='text-align:right'> </TD>
						<TD nowrap align=center><input name="FL" class="readonly8s" size=2 value="<%=tmpRec(CurrentPage, CurrentRow,8)%>" readonly style='text-align:right;<%if cdbl(tmpRec(CurrentPage, CurrentRow,8))<>0 then%>color:red<%end if%>'> </TD>
						<TD nowrap align=center><input name="FLmoney"  class="readonly8s" size="6"  value="<%=tmpRec(CurrentPage,CurrentRow,12+cdbl(icount))%>" readonly style='text-align:right' > </TD>
						<TD align=center><input name="FQD"  class="inputbox8r" size="5" value="0" onblur="FQDchg(<%=currentRow-1%>)" > </TD>
						<TD align=center>
							<input name="SSmoeny"  class="readonly8s" size="6" value="<%=tmpRec(CurrentPage,  CurrentRow,  9)%>" style='text-align:right' ONCHANGE=SUKMCHG(<%=currentRow-1%>)>
							<input TYPE=HIDDEN name="BSSmoeny"  class="readonly8s"   value="<%=tmpRec(CurrentPage,  CurrentRow,  9)%>" >							
						</TD> 
						<TD align=center>
							<%if tmpRec(CurrentPage,CurrentRow,22+cdbl(icount))< 70 then %>
								<input name="khfen"  class="readonly8s" readonly size="3" value="<%=tmpRec(CurrentPage,CurrentRow,22+cdbl(icount))%>"  style='text-align:right;color:red' > 
							<%else%>
								<input type=hidden name="khfen"   size="5" value="<%=tmpRec(CurrentPage,CurrentRow,22+cdbl(icount))%>" > 
							<%end if %>	
						</TD>						
						<TD  align=center>
							<input name="TOTJX" class="inputbox8r" size=10 value="<%=tmpRec(CurrentPage,CurrentRow,14+icount)%>" style='color:blue'>
							<input type=hidden name="workJs" class="inputbox8r" size=5 value="<%=tmpRec(CurrentPage,CurrentRow,16+icount)%>" >
							<input type=hidden name="BTOTJX"  value="<%=tmpRec(CurrentPage,CurrentRow, 14+icount)%>" >
							<input name="realJXM" class="inputbox8r" size=10 value="<%=tmpRec(CurrentPage,CurrentRow,17+icount)%>" type=hidden>
						</TD>						
					</TR>
				<%else%>
					<input type=hidden name="empid"  >
					<%for  zz  =  1  to  ICOUNT  %>
						<input type=hidden name="JX"&<%=arrays(zz,5)%>>
						<input type=hidden name="BJX"&<%=arrays(zz,5)%>>
					<%next%>

					<input type=hidden name="FL" >
					<input type=hidden name="FLmoney" >
					<input type=hidden name="FQD" >
					<input type=hidden name="SSmoeny" >
					<input type=hidden name="BSSmoeny">
					<input type=hidden name="TOTJX" >
					<input type=hidden name="BTOTJX" >	
					<input type=hidden name="workJs" >	
					<input type="hidden" name="realJXM" >	
			<%end if%>
			<%next%>
	</TABLE>
	<br>
	<table width=600 class=txt9>
		<tr>
		<td align=left>
		<% If CurrentPage > 1 Then %>
			<input type="submit"name="send"  value="FIRST" class=button>
			<input type="submit"  name="send" value="BACK" class=button>
		<% Else  %>
			<input  type="submit" name="send" value="FIRST"  disabled  class=button>
			<input  type="submit"  name="send" value="BACK" disabled class=button>
		<% End If %>
		<%  If  cint(CurrentPage)  <  cint(TotalPage)  Then %>
			<input type="submit" name="send" value="NEXT" class=button>
			<input type="submit" name="send" value="END" class=button>
		<% Else %>
			<input type="submit" name="send" value="NEXT" disabled class=button>
			<input type="submit" name="send" value="END"  disabled  class=button>
		<%  End  If  %>
		</TD>
		<td align=center>共<%=RecordInDB%>筆, 第<%=CurrentPage%>頁/共<%=TotalPage%>頁</td>
		<td  align=riht>
			<input type=button  name=btm  class=button value="確　　認" onclick="go()">
			<input Type=reset name=btM class=buttoN value="取　　消">
		</td>
		</tr>
	</table>
	<input type=hidden name="empid" >
	<%for yy = 1to ICOUNT %>
	<input type=hidden name="JX"&<%=arrays(yy,5)%>>
	<input type=hidden name="BJX"&<%=arrays(yy,5)%>>
	<%next%>
	<input type=hidden name="SUMJX" >
	<input type=hidden name="FL" >
	<input type=hidden name="FLmoney" >
	<input type=hidden name="FQD" >
	<input type=hidden name="SSmoeny" >
	<input type=hidden name="BSSmoeny" >
	<input type=hidden name="TOTJX" >
	<input type=hidden name="BTOTJX" >
	<input type=hidden name="workJs" >		
	<input type="hidden" name="realJXM" >	
</td></tr></table>
</form>
</body>
</html>
<script language=vbscript >
function oepnEmpWKT(index)  
	empidstr  =<%=self%>.empid(index).value 
	yymmstr  = <%=self%>.JXym.value 	 
	open "../ZZZ/getEmpWorkTime.asp?yymm="& yymmstr & "&empid=" & empidstr , "_blank","top=10 , left=10, scrollbars=yes"
end function

function  datachg(index)		 
	thiscols  =document.activeElement.name 
	Bcols="B"&document.activeElement.name
	if isnumeric(document.all(thiscols)(index).value)=false then
		alert  "請輸入數值"
		document.all(thiscols)(index).focus()
		document.all(thiscols)(index).value=document.all(Bcols)(index).value
		document.all(thiscols)(index).select()
	 	'exit function 		 
	else
		NewJXM=cdbl(document.all(Bcols)(index).value) -  cdbl(document.all(thiscols)(index).value)
		NewTOTJX=cdbl(<%=self%>.BTOTJX(index).value)-cdbl(<%=self%>.TOTJX(index).value)
		<%=self%>.TOTJX(index).value=cdbl(<%=self%>.TOTJX(index).value) - (cdbl(NewJXM)) - (+ewTOTJX)
	end if
end function

FUNCTION SUKMCHG(INDEX)
	IF ISNUMERIC(<%=SELF%>.SSMOENY(INDEX).VALUE)=FALSE THEN
		ALERT "請輸入數值!!"
		<%=SELF%>.SSMOENY(INDEX).FOCUS()
		<%=SELF%>.SSMOENY(INDEX).VALUE=<%=SELF%>
	ELSE
		NEWSUKM=CDBL(<%=SELF%>.BSSMOENY(INDEX).VALUE)-CDBL(<%=SELF%>.SSMOENY(INDEX).VALUE)
		NewTOTJX=cdbl(<%=self%>.BTOTJX(index).value)-cdbl(<%=self%>.TOTJX(index).value)
		<%=self%>.TOTJX(index).value=cdbl(<%=self%>.TOTJX(index).value) + (cdbl(NEWSUKM)) + (CDBL(NewTOTJX))
	END IF
END FUNCTION

function  fqdchg(index)  
	'200803反規定ㄧ次扣10000VND
	if <%=self%>.fqd(index).value<>"" then 
		if isnumeric(<%=self%>.fqd(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.fqd(index).value="0"
			<%=self%>.fqd(index).select()
			exit function 
		elseif cdbl(<%=self%>.fqd(index).value)=0 then
			<%=self%>.TOTJX(index).value = cdbl(<%=self%>.sumjx(index).value)- cdbl(<%=self%>.FLmoney(index).value)
		else
			<%=self%>.TOTJX(index).value = cdbl(<%=self%>.sumjx(index).value) - cdbl(<%=self%>.FLmoney(index).value)-( cdbl(<%=self%>.fqd(index).value)*10000 ) 
		end if 
	end if	 
	
	'old 適用至2008/02
	'if  cdbl(<%=self%>.fqd(index).value)=1  then
	'	<%=self%>.TOTJX(index).value=cdbl(<%=self%>.TOTJX(index).value)/2
	'elseif cdbl(<%=self%>.fqd(index).value)=2 then
	'	<%=self%>.TOTJX(index).value=0
	'else
 	'	<%=self%>.TOTJX(index).value = <%=self%>.TOTJX(index).value 
	'end if
	
end function

function TT(a) 
	alert <%=self%>.JXSTT(a).value
end function

FUNCTION GO() 

	<%=SELF%>.SUBMIT
END FUNCTION

</script> 