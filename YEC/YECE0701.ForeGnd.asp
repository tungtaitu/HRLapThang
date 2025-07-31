<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "YECE0701"

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
shift = request("shift")

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
'RESPONSE.WRITE days

'本月節假日天數
SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' AND  isnull(status,'')<>'H1' "
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
'RESPONSE.WRITE  MMDAYS &"<BR>"
'RESPONSE.END
'----------------------------------------------------------------------------------------

'RECALC = REQUEST("recalc")
'IF recalc="Y"  THEN 
'	sql="delete empbhgt where  yymm='"& yymm &"' "
'	conn.execute(Sql) 
'end if 	


gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 52    'number of fields per record 

'test = "BB+CV"  

sqln ="select * from empbh_set where w='"& whsno &"'  and ym='"& yymm &"' " 
set ors=conn.execute(Sqln)
if ors.eof then 
	response.write "本月保險計算未設定(CE/2/2.1)" 
	response.end 
else
	setstr =  left(ors("setstr"),len(ors("setstr"))-1) 
	c_cols=split(replace(ors("setstr"),"+",","),",")
	'response.write c_cols &"<BR>"
end if 
set ors=nothing   

allcols = ubound(c_cols) '欄位數
redim A1(allcols,2)
for k = 1 to ubound(c_cols) 
	showCols = showCols &  c_cols(k-1)&" as C"& k &","  	 
	A1(k,0)= c_cols(k-1) 
next  
TableRec = TableRec + cdbl(allcols)+5 
Session("a1cols") = A1 
'response.write allcols & "----" & showCols  
sql="select "  	
for  xx = 1 to allcols 
	sql=sql & "isnull(c.c"&xx &",0) as c"&xx &","
next 
'"left join (select * from empsalaryBasic )l on a.country=l.country and a.job=l.job "&_
sql=sql & " lncode, cast(isnull(c.BHT1,0) as decimal(18,0)) BHT1, isnull(c.bb,0) N_BB, isnull(c.empid,'') eid, isnull(c.cv,0) N_cv, isnull(c.phu,0) n_phu,  "&_
	"BHXH =  case when isnull(b.empid,'')='' then CASE WHEN CONVERT(CHAR(10), A.BHDAT, 111)<'"& ccdt &"' AND ISNULL(A.BHDAT,'')<>'' THEN isnull(c.bht1,0)*(isnull(f.emp_bhxh,0)/100.00) ELSE 0 END else isnull(b.bhxh5,isnull(c.bb,0)*isnull(f.emp_bhxh,0)/100.00) end  ,  "&_
	"BHYT =  case when isnull(b.empid,'')='' then CASE WHEN CONVERT(CHAR(10), A.BHDAT, 111)<'"& ccdt &"' AND ISNULL(A.BHDAT,'')<>'' THEN isnull(c.bht1,0)*(isnull(f.emp_bhyt,0)/100.00) ELSE 0 END else isnull(b.BHYT1,isnull(c.bb,0)*isnull(f.emp_bhyt,0)/100.00) end  ,   "&_
	"BHTN =  case when isnull(b.empid,'')='' then CASE WHEN CONVERT(CHAR(10), A.BHDAT, 111)<'"& ccdt &"' AND ISNULL(A.BHDAT,'')<>'' THEN isnull(c.bht1,0)*(isnull(f.emp_bhtn,0)/100.00) ELSE 0 END else isnull(b.BHTn1,isnull(c.bb,0)*isnull(f.emp_bhtn,0)/100.00) end  ,   "&_
	"GTAMT = case when isnull(b.empid,'')='' then case when ( isnull(a.gtdat,'')<>''AND isnull(a.gtdat,'')<='"& yymm &"' ) then isnull(f.emp_gtant,0) else 0 end else isnull(b.gtamt,0) end , "&_
	"flag=case when isnull(b.empid,'')='' then 'Y' else 'N' end , isnull(b.kh1,0) kh1, isnull(b.chanjia,0) chanjia, isnull(b.memo,'') bhmemo, "&_
	"isnull(d.HHOUR ,0 ) as jiaHr , isnull(e.HHOUR,0) as canJia, a.* from  "&_
	"( select * from  view_empfile  where  country='VN'  and CONVERT(CHAR(10), indat, 111)< '"& ccdt &"' and ( isnull(outdat,'')='' or outdat>'"& calcdt &"' )  ) a  "&_
	"left join ( select * from empbhgt  where yymm='"& YYMM &"'  ) b on b.empid = a.empid  "&_   
	"left join ( select "& showCols &" "& setstr &" as BHT1,  * from bemps where yymm='"& yymm &"'  ) c on c.empid = a.empid "&_ 	
	"left join ( SELECT EMPID, ISNULL(SUM(HHOUR),0) HHOUR  FROM EMPHOLIDAY WHERE  CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
	"AND JIATYPE in ('A','B' ) GROUP BY EMPID  )  d on d.empid = a.empid "&_	
	"left join ( SELECT EMPID, ISNULL(SUM(HHOUR),0) HHOUR  FROM EMPHOLIDAY WHERE  CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
	"AND JIATYPE ='F' GROUP BY EMPID  )  e on e.empid = a.empid "&_	   
	"left join (select * from empbh_per where yymm='"& yymm &"'  ) f on f.country=case when a.country='VN' then 'VN' else 'HW'  end  "&_ 	
	"where 1=1   "&_
	"and a.whsno like '%"& whsno &"%' and a.unitno like '%"& unitno &"%' and a.groupid like '%"& groupid &"%'  "&_
	"and a.COUNTRY like '%"& COUNTRY  &"%' and A.job like '%"& job &"%' and a.empid like '%"& QUERYX &"%' "
	if outemp="D" then
		sql=sql&" and ( isnull(a.outdat,'')<>'')  "
	elseif 	outemp="N" then
		sql=sql&" and ( isnull(a.outdat,'')='' )  "
	elseif outemp="TS" then 
		sql=sql&" and isnull(e.hhour,0)>="&cdbl(MMDAYS)*8&" " 
	end if
	if shift="C" then
		sql=sql&" and isnull(a.shift,'') NOT IN ('N', 'A', 'B' ) "
	ELSE
		sql=sql&" and isnull(a.shift,'') like '%"& shift &"'    "
	end if
	sql=sql&"order by a.empid   "
 
'response.write sql &"<P>"
'response.end
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
			tmpRec(i, j, 17)=RS("autoid")
			tmpRec(i, j, 18)=RS("outdate")
			tmpRec(i, j, 19)=RS("lncode")		 'code
			'tmpRec(i, j, 40) = rs("code")
			if rs("flag") = "N" then 	
				tmpRec(i, j, 20)=RS("n_bb")
			else
				tmpRec(i, j, 20)=RS("n_bb")
			end if 	
			tmpRec(i, j, 21)=RS("BHXH")
			tmpRec(i, j, 22)=RS("BHYT")
			tmpRec(i, j, 23)=RS("GTAMT")			
			tmpRec(i, j, 24)=RS("flag")	
			if rs("flag")="N" then 
				tmpRec(i, j, 25)="Blue"
			else
				tmpRec(i, j, 25)="Black"
			end if 
			
			tmpRec(i, j, 26)=RS("BHDAT")
			tmpRec(i, j, 27)=RS("GTDAT")
			'tmpRec(i, j, 28)=CDBL(tmpRec(i, j, 21))+CDBL(tmpRec(i, j, 22))
			tmpRec(i, j, 29)=rs("KH1")
			tmpRec(i, j, 30)=rs("chanjia")
			tmpRec(i, j, 31)=rs("bhmemo")
			tmpRec(i, j, 32)=rs("BHTN")  '失業保險 since 200901 
			
			tmpRec(i, j, 28)=CDBL(tmpRec(i, j, 21))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 32))
			'response.write   tmpRec(i, j, 20) &"<Br>"
			'response.write  
			if rs("eid")="" then 
				tmpRec(i, j, 33)="*"
			else
				tmpRec(i, j, 33)=""
			end if 	
			tmpRec(i, j, 34) = rs("bht1")
			if rs("flag") = "N" then 	
				tmpRec(i, j, 35) = rs("N_CV")
			else
				tmpRec(i, j, 35) = rs("N_CV")
			end if 	 			
			tmpRec(i, j, 36) = rs("n_phu")  
			s1=""
			
			for yy = 1 to allcols  
				colsname="C"&yy 
				'response.write  rs(colsname)   
				tmpRec(i, j, 36+yy) =  rs(colsname)   
				'response.write "XXX=" & 36+yy & " " & tmpRec(i, j, 36+yy) &"<BR>"
				s1 = s1& A1(yy,0)&"+"			
			next 
			
			tmpRec(i, j, 36+allcols+1) = left(s1,len(s1)-1)
			tmpRec(i, j, 37+allcols+1) = rs("jiaHr")  
			'response.write   rs("jiaHr")  &"<BR>" 
			if ( cdbl(rs("jiaHr"))/8 >= cdbl(mmdays) or cdbl(rs("jiaHr"))>=cdbl(MMDAYS*8) ) then 
				'response.write "xxx" &"<BR>" 
				'response.write cdbl(rs("jiaHr"))/8.0 &"<BR>"
				tmpRec(i, j, 21)=0
				tmpRec(i, j, 22)=0
				tmpRec(i, j, 32)=0  '失業保險 since 200901 
				tmpRec(i, j, 28)=CDBL(tmpRec(i, j, 21))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 32))
				tmpRec(i, j, 33)=tmpRec(i, j, 33) & "本月請(事、病、產假)"
			end if 
			
			'response.write cdbl(rs("canJia")) &  cdbl(MMDAYS*8) & "<BR>"
			if cdbl(rs("canJia"))>=cdbl(MMDAYS/2*8) and rs("flag")="Y" then 	'本月請產假			
				'2011/12/05
				'modify:Steven
				'requester:Thúy
				'tmpRec(i, j, 30) = 1
				tmpRec(i, j, 21)=0
				tmpRec(i, j, 22)=0
				tmpRec(i, j, 32)=0  '失業保險 since 200901 
				'tmpRec(i, j, 28)=CDBL(tmpRec(i, j, 21))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 32))
				'公司先付23% , 待員工上班後, 保險局繳回公司20% 2009/09/30  
				'response.write "XXX"&"<BR>"
				tmpRec(i, j, 28) = CDBL(tmpRec(i, j, 21))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 32))
				tmpRec(i, j, 33)=tmpRec(i, j, 33) & "本月請產假"
				
			end if 	
			
			'response.write   RS("BHTN") &"<Br>"
			'response.end
			
			tmpRec(i, j, 50)=RS("BHXH")
			tmpRec(i, j, 51)=RS("BHYT")
			tmpRec(i, j, 52)=rs("BHTN")  '失業保險 since 200901 
			
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
	Session("empBHGTD") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empBHGTD")

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
<body  leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()" bgproperties="fixed"  >
<form name="<%=self%>" method="post" action="<%=SELF%>.ForeGnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>">
<INPUT TYPE=hidden NAME=whsno VALUE="<%=whsno%>">
<INPUT TYPE=hidden NAME="allcols" value="<%=allcols%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3> 
								<tr CLASS="txt">
									<Td><font color=Red>計薪年月(Thang Nam)：<%=YYMM%></font></td>
									<td><b>memo=* : khong co lương cơ bản(尚未設定基本薪)</b></td>
								</tr>	
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%">
								<TR HEIGHT=25 BGCOLOR="#C0C0C0"  class=txt8 >
									<TD WIDTH=40 ALIGN=CENTER nowrap>項次<BR>STT</TD>
									<TD WIDTH=60 align=center>工號<BR>So The</TD>
									<TD WIDTH=120 NOWRAP >員工姓名(中,英,越)<BR>Ho Ten</TD> 		
									<td WIDTH=55 align=center nowrap>到職日期<BR>NVX(yy/mm/dd)</td>
									<td WIDTH=55 align=center nowrap>離職日期<BR>NTV(yy/mm/dd)</td> 
									 <td WIDTH=55 align=center nowrap>Bac Luong</td> 		
									<%for kk = 1 to allcols 
										s1 = s1& A1(kk,0)&"+"
										Select case A1(kk,0) 
											Case "BB"
												 myText = "基薪<br>Cơ bản"
											Case "CV"
												 myText = "職務加給<br>Chức vụ"
											Case "KT"
												 myText = "技術<br>Kỷ thuật"
											Case "MT"
												 myText = "環境<br>Môi trường"
											Case "NN"
												 myText = "燃油津貼<br>PC xăng xe"
											Case "PHU"
												 myText = "電話津貼<br>PC điện thoại"
											Case Else 
												 myText A1(kk,0) 	
										end Select
									%>
										<TD align=center nowrap><%=myText%></TD>			
									<%next%> 		
									<TD align=center nowrap>保險費<br>PBH</TD>
									<td WIDTH=55 align=center>保險日期<BR>NBH(../dd)</td>
									<td WIDTH=50 nowrap align=center>PHÁT<Br>SINH<Br>Thang</td>
									<td WIDTH=50 nowrap align=center>THAI<Br>SẢN<Br>Thang</td>
									<TD WIDTH=40 nowrap align=center>BHXH</TD>
									<TD WIDTH=40 nowrap align=center>BHYT</TD>
									<TD WIDTH=40 nowrap align=center>BHTN</TD>
									<TD WIDTH=60 nowrap align=center>BHTOT</TD>
									<td WIDTH=50 nowrap align=center>入工團</td>
									<TD WIDTH=40 nowrap align=center>工團費</TD>
									<TD WIDTH=100 nowrap align=center>memo</TD>
								</tr>
								<%for CurrentRow = 1 to PageRec
									IF CurrentRow MOD 2 = 0 THEN
										WKCOLOR="LavenderBlush"
										'wkcolor="#ffffff"
									ELSE
										'WKCOLOR="#DFEFFF"
										'WKCOLOR="#D7EBFF"
										WKCOLOR="#E8F3FF" 			
										'wkcolor="#ffffff"
									END IF
									'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
								%>
								<TR BGCOLOR=<%=WKCOLOR%> >
									<TD ALIGN=CENTER ><FONT COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>">
									<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*15)%><%END IF %></FONT>
									</TD>
									<TD ALIGN=CENTER>  		
										<%=tmpRec(CurrentPage, CurrentRow, 1)%>	
										<!--a href="vbscript:view1(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)">
											<FONT COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>"><%=tmpRec(CurrentPage, CurrentRow, 1)%></FONT>
										</a-->
										<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
										<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
									</TD>
									<TD  >
										<a href='vbscript:editmemo(<%=CurrentRow-1%>)'>
											<FONT COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>"><%=tmpRec(CurrentPage, CurrentRow, 2)%></FONT><br>
											<font COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
										</a>
									</TD> 		
									<TD  ALIGN=CENTER nowrap width=55><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD>
									<TD  ALIGN=CENTER nowrap width=55><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 18),8)%></FONT></TD> 		
									<TD  ALIGN=CENTER nowrap width=55><FONT CLASS=TXT8><%=tmpRec(CurrentPage, CurrentRow, 19) %></FONT></TD> 
									<%for zz = 1 to allcols %>
									<TD ALIGN=RIGHT>
										<%if tmpRec(CurrentPage, CurrentRow, 1) <> "" then%> 
											<input type="text" name="<%=a1(zz,0)%>"  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 36+zz)))%>"  readonly  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  >							
										<%else%>
											<input type=hidden size=5 name="<%=a1(zz,0)%>"  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 36+zz)))%>">			
										<%end if%>			
									</TD>
									<%next%>  
									<TD ALIGN=RIGHT>
										<%if tmpRec(CurrentPage, CurrentRow, 1) <> "" then%> 
											<input type="text" name="BHP"  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 34)))%>"  readonly  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" >			
										<%else%>
											<input type=hidden size=5 name="BHP"  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 34)))%>">			
										<%end if%>			
										<input type=hidden size=5 name="BBCODE"  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 19)))%>">			
									</TD>			 
									<TD  ALIGN=CENTER nowrap width=55><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 26),8)%></FONT></TD>
									<TD ALIGN=CENTER >
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
											<input type="text" STYLE="width:100%;" name=KH1   value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 29)))%>" ONCHANGE="DATACHG(<%=currentrow-1%>)">
										<%else%>	
											<input name=KH1  size=5 type=hidden>
										<%end if %>
									</TD>
									<TD ALIGN=CENTER >
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
											<input type="text" STYLE="width:100%;" name=chanjia value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 30)))%>" ONCHANGE="DATACHG(<%=currentrow-1%>)" >
										<%else%>	
											<input name=chanjia  size=5 type=hidden >
										<%end if%>
									</TD> 		
									<TD ALIGN=RIGHT>
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
											<INPUT type="text" NAME="BHXH"  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>"  STYLE="width:100%;TEXT-ALIGN:RIGHT"  ONCHANGE="DATACHG1(<%=currentrow-1%>)" > 
										<%else%>
											<INPUT NAME="BHXH" TYPE='HIDDEN'>-
										<%end if%> 		
									</TD>
									<TD ALIGN=RIGHT>
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
											<INPUT type="text" NAME="BHYT"  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>"  STYLE="width:100%;TEXT-ALIGN:RIGHT"  ONCHANGE="DATACHG(<%=currentrow-1%>)" > 
										<%else%>
											<INPUT NAME="BHYT" TYPE='HIDDEN'>-
										<%end if%>
									</TD>
									<TD ALIGN=RIGHT>
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
											<INPUT type="text" NAME="BHTN" VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>"  STYLE="width:100%;TEXT-ALIGN:RIGHT"  ONCHANGE="DATACHG(<%=currentrow-1%>)" > 
										<%else%>
											<INPUT NAME="BHTN" TYPE='HIDDEN'>-
										<%end if%>
									</TD>		
									<TD ALIGN=RIGHT>
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
											<INPUT type="text" NAME="BHTOT"  READONLY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 28)%>"  STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  > 
										<%else%>
											<INPUT NAME="BHTOT" TYPE='HIDDEN'>-
										<%end if%>
									</TD> 
									<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=tmpRec(CurrentPage, CurrentRow, 27)%></FONT></TD>
									<TD ALIGN=RIGHT>
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
											<INPUT type="text" NAME="GTAMT"  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>"  STYLE="width:100%;TEXT-ALIGN:RIGHT"  ONBLUR="DATACHG(<%=currentrow-1%>)"  > 
										<%else%>
											<INPUT NAME="GTAMT" TYPE='HIDDEN'>-
										<%end if%>
									</TD>
									<td align="left"><font color="red"><%=tmpRec(CurrentPage, CurrentRow, 33)%>&nbsp;</font></td>
								</TR>
								<%next%>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td>
							<input type=hidden name=empid>
							<input type=hidden name=BBCODE>
							<%for z1=1 to allcols%>
								<input type=hidden name="<%=a1(z1,0)%>">
							<%next%>
							<input type=hidden name=BHP>
							<input type=hidden name=kh1>
							<input type=hidden name=chanjia>
							<INPUT NAME="BHXH" TYPE='HIDDEN'>
							<INPUT NAME="BHYT" TYPE='HIDDEN'>
							<INPUT NAME="BHTN" TYPE='HIDDEN'>
							<INPUT NAME="GTAMT" TYPE='HIDDEN'>
							<INPUT NAME="BHTOT" TYPE='HIDDEN'>

							<table class="table-borderless table-sm bg-white text-secondary">
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
									<TD WIDTH=25% ALIGN=RIGHT nowrap>
										<input type="BUTTON" name="send" value="(Y)COnfirm" class="btn btn-sm btn-danger" ONCLICK="GO()">
										<input type="BUTTON" name="send" value="(N)Cancel" class="btn btn-sm btn-outline-secondary" onclick="clr()">
										<input type="BUTTON" name="send" value="(R)重新計算" class="btn btn-sm btn-outline-secondary" onclick="clrb()">
									</TD>
								</TR>
							</TABLE>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>




</form>




</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("empBHGTD")
	for CurrentRow = 1 to PageRec
		'tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 19) = request("BBCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 20) = request("BB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 21) = request("BHXH")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 22) = request("BHYT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("GTAMT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 28) = request("BHTOT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 29) = request("kh1")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 30) = request("chanjia")(CurrentRow)
	
	next
	Session("empBHGTD") = tmpRec

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
	open "<%=self%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&code=" &	codestr , "Back"	
	'PARENT.BEST.COLS="70%,30%"
END FUNCTION


FUNCTION bbchg(INDEX)
	codestr=<%=self%>.bb(index).value	
	<%=self%>.BHXH(index).value = cdbl(codestr)*0.05
	<%=self%>.BHYT(index).value = cdbl(codestr)*0.01
	<%=self%>.BHTOT(index).value = cdbl(<%=self%>.BHXH(index).value)+cdbl(<%=self%>.BHYT(index).value)
	'open "empbhgt.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
	'	 "&code=" &	codestr , "Back"	
	'PARENT.BEST.COLS="70%,30%"
END FUNCTION



FUNCTION DATACHG(INDEX)
	if isnumeric(<%=SELF%>.BHXH(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.BHXH(index).focus()
		<%=self%>.BHXH(index).value=0
		<%=self%>.BHXH(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.BHYT(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.BHYT(index).focus()
		<%=self%>.BHYT(index).value=0		
		<%=self%>.BHYT(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.BHTN(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.BHTN(index).focus()
		<%=self%>.BHTN(index).value=0		
		<%=self%>.BHTN(index).select()
		exit FUNCTION
	end if	
	if isnumeric(<%=SELF%>.GTAMT(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.GTAMT(index).focus()
		<%=self%>.GTAMT(index).value=0		
		<%=self%>.GTAMT(index).select()
		exit FUNCTION
	end if  	 
	
	if isnumeric(<%=SELF%>.kh1(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.kh1(index).focus()
		<%=self%>.kh1(index).value=0		
		<%=self%>.kh1(index).select()
		exit FUNCTION
	end if  	 

	if isnumeric(<%=SELF%>.chanjia(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.chanjia(index).focus()
		<%=self%>.chanjia(index).value=0		
		<%=self%>.chanjia(index).select()
		exit FUNCTION
	end if  	 	
	
	<%=SELF%>.BHTOT(INDEX).VALUE=CDBL(<%=self%>.BHXH(index).value)+CDBL(<%=self%>.BHYT(index).value)+CDBL(<%=self%>.BHTN(index).value)

	CODESTR01 = <%=SELF%>.BB(INDEX).VALUE
	CODESTR02 = <%=SELF%>.BBCODE(INDEX).VALUE
	CODESTR03 = <%=SELF%>.BHXH(INDEX).VALUE
	CODESTR04 = <%=SELF%>.BHYT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.GTAMT(INDEX).VALUE   
	CODESTR06 = <%=SELF%>.kh1(INDEX).VALUE   
	CODESTR07 = <%=SELF%>.chanjia(INDEX).VALUE   
	CODESTR08 = <%=SELF%>.BHTN(INDEX).VALUE   
	

	open "<%=self%>.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 &_
		 "&CODESTR06="& CODESTR06 &_
		 "&CODESTR07="& CODESTR07 &_
		 "&CODESTR08="& CODESTR08  , "Back"

	'PARENT.BEST.COLS="70%,30%"

END FUNCTION  


FUNCTION DATACHG1(INDEX)
	if isnumeric(<%=SELF%>.BHXH(INDEX).VALUE)=false then
		alert "請輸入數字!!"
		<%=self%>.BHXH(index).focus()
		<%=self%>.BHXH(index).value=0
		<%=self%>.BHXH(index).select()
		exit FUNCTION
	end if  

	CODESTR01 = <%=SELF%>.BB(INDEX).VALUE
	CODESTR02 = <%=SELF%>.BBCODE(INDEX).VALUE
	CODESTR03 = <%=SELF%>.BHXH(INDEX).VALUE
	CODESTR04 = <%=SELF%>.BHYT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.GTAMT(INDEX).VALUE   
	CODESTR08 = <%=SELF%>.BHTN(INDEX).VALUE   
 
	IF <%=SELF%>.BHXH(INDEX).VALUE="0" THEN 
		<%=SELF%>.BHYT(INDEX).VALUE="0"
		CODESTR04 = "0"
		<%=SELF%>.BHTN(INDEX).VALUE="0"   
		CODESTR08 = "0" 
	END IF 
	<%=SELF%>.BHTOT(INDEX).VALUE=CDBL(<%=self%>.BHXH(index).value)+CDBL(<%=self%>.BHYT(index).value)
	
	open "<%=self%>.back.asp?ftype=CDATACHG1&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 &_
		 "&CODESTR08="& CODESTR08  , "Back"
	'PARENT.BEST.COLS="70%,30%"

END FUNCTION  

function view1(index)
	yymmstr = <%=self%>.yymm.value
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	OPEN "../zzz/getempWorkTime.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10 , scrollbars=yes"
end function


function editmemo(index)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	YYMM = <%=self%>.YYMM.value
	open "<%=self%>.memo.asp?index="& index &"&currentpage=" & cp &"&yymm=" & yymm  , "_blank" , "top=10, left=10, width=550,height=450, scrollbars=yes"
end function  


</script>

