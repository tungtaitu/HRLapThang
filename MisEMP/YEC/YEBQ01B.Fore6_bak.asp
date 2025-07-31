<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

SELF = "YEBQ01B"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

firstday = year(date())&"/"&right("00"&month(date()),2)&"/01" 
nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 
  
empautoid = TRIM(REQUEST("empautoid"))  
  
SQL="SELECT * FROM  view_empfile where ISNULL(STATUS,'')<>'D' AND  autoid='"& empautoid &"' "

'RESPONSE.WRITE SQL
'RESPONSE.END
RS.OPEN SQL , CONN, 3, 3
IF NOT RS.EOF THEN
	empautoid = TRIM(RS("AUTOID"))
	emptype = TRIM(RS("emptype"))
	EMPID=TRIM(RS("EMPID"))	'員工編號
	INDAT=TRIM(RS("nindat"))	'到職日
	WHSNO=TRIM(RS("WHSNO"))	'廠別
	UNITNO=TRIM(RS("UNITNO"))	'處/所
	GROUPID=TRIM(RS("GROUPID"))	'組/部門
	ZUNO=TRIM(RS("ZUNO"))	'單位

	EMPNAM_CN=TRIM(RS("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RS("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RS("COUNTRY"))	'國籍
	COUNTRYDESC=TRIM(RS("cstr"))	'國籍
	BYY=(TRIM(RS("BYY"))) '年(生日)
	BMM=(RS("BMM"))	'月(生日)
	BDD=(RS("BDD"))	'日(生日)
	AGES=TRIM(RS("AGES"))	'年齡
	SEX=TRIM(RS("SEX"))	'性別
	JOB=TRIM(RS("JOB"))  '職等 
	Jstr=trim(rs("jstr"))
	Gstr=trim(rs("gstr"))
	zstr=trim(rs("zstr"))
	wstr=trim(rs("wstr"))
	ustr=trim(rs("ustr"))
	PERSONID=TRIM(RS("PERSONID"))	'身分証字號
	BHDAT=TRIM(RS("BHDAT"))	'簽約日(保險日)
	PASSPORTNO=TRIM(RS("PASSPORTNO"))	'護照號碼
	VISANO=TRIM(RS("VISANO"))	'簽證號碼
	PISSUEDATE=TRIM(RS("PISSUEDATE")) '護照簽發日
	PDUEDATE=TRIM(RS("PDUEDATE"))	'護照有效期
	VDUEDATE=TRIM(RS("VDUEDATE"))	'簽證有效期	
	PHONE=TRIM(RS("PHONE"))	'聯絡電話
	MOBILEPHONE=TRIM(RS("MOBILEPHONE"))	'手機
	HOMEADDR=TRIM(RS("HOMEADDR"))	'聯絡地址
	EMAIL=TRIM(RS("EMAIL"))	'EMAIL
	OUTDATe=TRIM(RS("OUTDATe"))	'離職日
	MEMO=TRIM(RS("MEMO"))	'其他說明
	GTDAT=TRIM(RS("GTDAT"))	'加入工團(年月)
	MARRYED = trim(RS("MARRYED"))    '婚姻狀況
	SCHOOL=RS("SCHOOL") '教育程度
	SHIFT=RS("SHIFT") '班別
	grps = rs("grps") 
	studyjob=RS("studyjob") '職能學習 
	
	'PHOTOS=TRIM(RS("PHOTOS"))	'照片檔名 
	PHOTOS=RS("EMPID")&".JPG"
	'-----------------------------------------
	'PHU = RS("PHU")
	'NN = RS("NN")
	'KT = RS("KT")
	'TTKH = RS("TTKH")
	'MT = RS("MT")
	'BB = RS("BB")
	BANKID = RS("BANKID") 
	
	wkd_no = RS("wkd_no") 
	wkd_duedate = RS("wkd_duedate") 
	experience = RS("experience")
	urgent_person = RS("urgent_person")
	releation = RS("releation")
	urgent_addr = RS("urgent_addr")
	urgent_tel = RS("urgent_tel")
	urgent_mobile=RS("urgent_mobile")
	bh_person=RS("bh_person")
	bh_personID=RS("bh_personID")
	
	
END IF
SET RS=NOTHING 



FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<link rel="stylesheet" href="../Include/bar.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	'<%=self%>.indat.focus()
	'<%=self%>.indat.select() 
	<%=self%>.send(0).style.display=""
	<%=self%>.send(1).style.display=""
	'<%=self%>.send(2).style.display="" 
	<%=self%>.emptype.focus()
end function

function groupchg()
	code = <%=self%>.GROUPID.value
	open "<%=self%>.back.asp?ftype=groupchg&code="&code , "Back"
	'parent.best.cols="50%,50%"
end function

function unitchg()
	code = <%=self%>.unitno.value
	open "<%=self%>.back.asp?ftype=UNITCHG&code="&code , "Back"	
	'parent.best.cols="50%,50%"
end function
-->
</SCRIPT>  
</HEAD> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"  >
<form  name="<%=self%>" method="post" action="YEBE0102.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<input name=act value="EMPEDIT" type=hidden >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		員工個人資料查詢
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 
<table width="650" border="0" cellspacing="0" cellpadding="0" align=center >
	<tr><td nowrap>
		<div id="navcontainer"  >
			<ul id="navlist">
			<li><a href="vbscript:chgpage(1)">基本資料<BR>Tu lieu co ban<BR>&nbsp;</a></li>
			<li ><a href="vbscript:chgpage(2)">教育訓練/証執照<br>huan luyen/<BR>bang cap</a></li>			
			<li ><a href="vbscript:chgpage(4)">獎懲紀錄<BR>Tu lieu<BR>thuong phat</a></li>
			<li ><a href="vbscript:chgpage(5)">部門/晉升紀錄<BR>Nang chuc/<BR>don vi </a></li>
			<%if session("rights")<="1" or session("rights")="5" then %>
				<li id=active><a href="vbscript:chgpage(6)">薪資資料<BR>Tien luong<BR>&nbsp;</a></li>
			<%else%>
				<li id=active><a >薪資資料<BR>Tien luong<BR>&nbsp;</a></li>
			<%end if%>	
			</ul>
		</div> 
		</td>
	</tr>  
</table>    
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 

<TABLE WIDTH=550 CLASS=TXT BORDER=0 cellspacing="1" cellpadding="1" > 		 
	<TR  >
		<TD WIDTH=70 ALIGN=RIGHT >Số thẻ：</TD>
		<TD colspan=3  ><%=EMPID%> &nbsp;
		<%=EMPNAM_CN%>&nbsp;<%=EMPNAM_VN%> &nbsp;
		(<%=COUNTRYDESC%>)&nbsp;&nbsp;
		<input name=empid value=<%=empid%> type=hidden >
		<input name=empautoid value=<%=empautoid%> type=hidden >
		<input name=country value=<%=country%> type=hidden >				
	</tr>
	<tr>	
		<TD  ALIGN=RIGHT>NVX：</TD>
		<TD width=100><%=INDAT%>&nbsp;&nbsp;	</td>
		<td align="right">Đơn vị：
		<td> (<%=WSTR%>) <%=groupid%> <%=GSTR%>-<%=zuno%> <%=ZSTR%></TD>		
	</TR>	
	<tr>
		<TD ALIGN=RIGHT>Chuc Vu：</TD>
		<TD colspan=4><%=job%>&nbsp;<%=jstr%></TD>
	</tr>
</TABLE> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=600> 
<table width=550><tr><td align=center>
<TABLE WIDTH=665 CLASS=txt8 BORDER=0  cellspacing="1" cellpadding="2" bgcolor=black>	
	<TR bgcolor=#e4e4e4 height=22> 
		<Td  width=60 nowrap align=center >統計年月<br>Nam Thang</td>
		<Td  width=30 nowrap align=center >幣別</td>
		<Td  width=60 nowrap align=center >BB</td>
		<Td  width=60 nowrap align=center >CV</td>
		<Td  width=60 nowrap align=center  >Y(phu)</td>		
		<Td  width=60 nowrap align=center  >NN</td>		
		<Td  width=60 nowrap  align=center  >KT</td>
		<Td  width=60 nowrap  align=center  >MT</td>
		<Td  width=60 nowrap  align=center  >TTKH</td>
		<Td  width=60 nowrap  align=center  >全勤</td>
		<Td  width=60 nowrap  align=center  >合計</td>
		
		
	</tr>
	<%
	sqlt="select  isnull(c.bonus,0) as basicQc , a.bb+a.cv+a.phu+a.nn+a.kt+a.mt+a.ttkh+a.qc as tot, a.* from "&_
		 "( select * from bemps where empid='"& empid &"' ) a  "&_
		 "left join ( select * from view_empjob   ) b on b.empid = a.empid and b.lj = a.job and b.yymm = a.yymm "&_
		 "left join ( select * from view_empgroup   ) d on d.empid = a.empid and  d.yymm = a.yymm   "&_
		 "left join ( select *from empSalarybasic  where  func='CC' )  c  on c.bwhsno = d.lw and c.country = a.country and c.job = b.lj "&_
		 "order by a.yymm " 
 	'response.write sqlt
	set rds=conn.execute(sqlt)
	x = 0 
	while not rds.eof 
	x = x + 1 
	if x mod 2 = 0 then 
		wkcolor="lightyellow"
	else
		wkcolor="#ffffff"
	end if 
	 
	%>
		<Tr bgcolor="<%=wkcolor%>" class=txt8>
			<td align=center><%=rds("yymm")%></td>
			<td align=center><%=rds("dm")%></td>
			<td align=right><%=formatnumber(rds("bb"),0)%></td>
			<td align=right><%=formatnumber(rds("cv"),0)%></td>
			<td align=right><%=formatnumber(rds("phu"),0)%></td>
			<td align=right><%=formatnumber(rds("nn"),0)%></td>
			<td align=right><%=formatnumber(rds("kt"),0)%></td>
			<td align=right><%=formatnumber(rds("mt"),0)%></td>
			<td align=right><%=formatnumber(rds("ttkh"),0)%></td>
			<td align=right><%=formatnumber(rds("qc"),0)%></td>
			<td align=right><%=formatnumber(rds("tot"),0)%></td>
		</Tr>
	<%
	rds.movenext
	wend 
	%>
	<%set rst=nothing%>
</table>
</td></tr></table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600> 
 
<TABLE WIDTH=550>
		<tr ALIGN=center>
			<TD >
			<input type=button name=send value="關閉視窗(Close)"  class=button onclick="window.close">
			</TD> 
		</TR>
</TABLE>


</form>


</body>
</html>
		
<script language=vbscript>
<!-- 
function prt()
	'<%=self%>.action="YEBE0301.toexcel.asp"
	'<%=self%>.submit()
	'<%=self%>.send(0).style.display="none"
	'<%=self%>.send(1).style.display="none"
	'<%=self%>.send(2).style.display="none"
	'window.print()
	'<%=self%>.send(0).style.display=""
	'<%=self%>.send(1).style.display=""
	'<%=self%>.send(2).style.display=""	
end function  

function chkempid()
	if <%=self%>.country.value<>"" and <%=self%>.whsno.value<>"" then 		
		code1=ucase(trim(<%=self%>.country.value))
		code2=ucase(trim(<%=self%>.whsno.value))
		code3=ucase(trim(<%=self%>.emptype.value)) 
		'alert  code1
		open "<%=self%>.back.asp?ftype=getempid&code1=" & code1 &"&code2=" & code2 &"&code3=" & code3 , "Back" 
		'parent.best.cols="70%,30%"
	end if 
end function 

function empidchg()
	empidstr = Ucase(Trim(<%=self%>.empid.value))
	if empidstr<>"" then
		open "<%=self%>.back.asp?ftype=empidchk&code="& empidstr , "Back"
		'parent.best.cols="50%,50%"
	end if
end function

function sexchg(x)
	if <%=self%>.radio1(0).checked=true then
		<%=self%>.sexstr.value="M"
	elseif 	<%=self%>.radio1(1).checked=true then
		<%=self%>.sexstr.value="F"
	else
		<%=self%>.sexstr.value=""
	end if
end function 

function typechg(x)
	if <%=self%>.radio3(0).checked=true then
		<%=self%>.emptype.value="A"
	elseif 	<%=self%>.radio3(1).checked=true then
		<%=self%>.emptype.value="B"
	elseif 	<%=self%>.radio3(2).checked=true then
		<%=self%>.emptype.value="C"
	else	
		<%=self%>.emptype.value=""
	end if 
	 
end function

function marrychg(x)
	if <%=self%>.radio2(0).checked=true then
		<%=self%>.marryed.value="Y"
	elseif 	<%=self%>.radio2(1).checked=true then
		<%=self%>.marryed.value="N"
	elseif 	<%=self%>.radio2(2).checked=true then
		<%=self%>.marryed.value="L"	
	else
		<%=self%>.marryed.value=""
	end if
end function

function BACKMAIN()
	open "../main.asp" , "_self"
end function

FUNCTION GO()
	IF  <%=SELF%>.EMPID.VALUE="" THEN
		ALERT "請輸入員工編號!!"
		<%=SELF%>.EMPID.FOCUS()
		EXIT FUNCTION 
	END IF
	'if <%=self%>.unitno.value="" then 
	'	ALERT "請輸入處/所!!"
	'	<%=SELF%>.unitno.FOCUS()
	'	EXIT FUNCTION 
	'end if 
	'if <%=self%>.GROUPID.value="" then 
	'	ALERT "請輸入部門單位!!"
	'	<%=SELF%>.GROUPID.FOCUS()
	'	EXIT FUNCTION 
	'end if 
	'if <%=self%>.shift.value="" then 
	'	ALERT "請輸入班別!!"
	'	<%=SELF%>.shift.FOCUS()
	'	EXIT FUNCTION 
	'end if 
	photosname=<%=self%>.empid.value&".jpg"
	<%=SELF%>.ACTION="YEBE0102.upd.asp"
	<%=SELF%>.SUBMIT
END FUNCTION

'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.pissueDate.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.pduedate.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.WKD_DueDate.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.outdat.value)
elseif a=6 then
	INcardat = Trim(<%=self%>.bhdat.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat.value=ANS
		elseif a=2 then
			Document.<%=self%>.pissueDate.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.WKD_DueDate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		elseif a=6 then
			Document.<%=self%>.bhdat.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.pissueDate.value=""
			Document.<%=self%>.pissueDate.focus()
		elseif a=3 then
			Document.<%=self%>.pduedate.value=""
			Document.<%=self%>.pduedate.focus()
		elseif a=4 then
			Document.<%=self%>.WKD_DueDate.value=""
			Document.<%=self%>.WKD_DueDate.focus()
		elseif a=5 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
		elseif a=6 then
			Document.<%=self%>.bhdat.value=""
			Document.<%=self%>.bhdat.focus()
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
		ELSE
			<%=SELF%>.AGES.VALUE=CDBL(YEAR(DATE()))-CDBL(<%=SELF%>.BYY.VALUE) + 1
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

function chgpage(a)
	code1=<%=self%>.empautoid.value
	code2=<%=self%>.empid.value
	if a = 1 then 
		if <%=self%>.country.value="VN" then 
			open "<%=self%>.editvn.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
		else
			open "<%=self%>.editHW.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
		end if	
	elseif a=2 then 
		open "<%=self%>.Fore3.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=3 then 
		open "<%=self%>.Fore3.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=4 then 
		open "<%=self%>.Fore4.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=5 then 
		open "<%=self%>.Fore5.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=6 then 
		open "<%=self%>.Fore6.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	else
	end if 
	
end function 
-->
</script>

