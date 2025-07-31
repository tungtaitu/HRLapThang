<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next   
session.codepage="65001"
SELF = "empholiday"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   


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
	<%=self%>.EMPID.focus()	
end function   

function datachg()
	<%=self%>.action="empwork.fore.asp?totalpage=0"
	<%=self%>.submit
end function  

function empidchg()
	if <%=self%>.empid.value<>"" then 
		<%=self%>.empid.value=Ucase(<%=self%>.empid.value) 
		codestr = Ucase(Trim(<%=self%>.empid.value)) 
		'alert codestr
		open "<%=self%>.back.asp?ftype=chkempid&code=" & codestr , "Back"  
		'parent.best.cols="50%,50%"
	end if
end function 

-->
</SCRIPT>  
</head>   
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.InsertDB.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">
							<table id="myTableForm" width="50%"> 
								<tr><td height="35px" colspan=4>&nbsp;</td></tr>  
								<TR>
									<TD nowrap align=right>員工編號<BR><FONT class="txt8">So the</font></TD>
									<TD COLSPAN=3 nowrap>
										<INPUT type="text" style="width:15%" NAME=EMPID  onchange="empidchg()">
										<INPUT type="text" style="width:25%" NAME=EMPNAMEVN  READONLY  >
										<INPUT type="text" style="width:25%" NAME="NVx"  READONLY  >												
										<INPUT type="text" style="width:23%" NAME="ntv" READONLY  >
										<INPUT NAME="country" READONLY type="hidden" >																					
									</TD> 
								</TR>								
								<TR>
									<TD align=right >尚有年假<BR><FONT class="txt8">Gio PN</font></TD>
									<TD COLSPAN=3 >
										<input type="text" style="width:25%;text-align:center" name="ynjh"  readonly value="">
										<input type="text" style="width:70%" name="ynj"  readonly value="">
									</TD>
								</TR>	
								<TR>
									<TD align=right >請假類別<BR><FONT class="txt8">Loai phep</font></TD>
									<TD COLSPAN=3 nowrap>										
										<SELECT NAME=HOLIDAY_TYPE  onchange="jiatypechg()" style="width:120px"> 
											<%sql="select * from basicCode where func='JB' order by sys_type" 
											set rds=conn.execute(sql)
											while not rds.eof 
											%>
											<OPTION value="<%=rds("sys_type")%>"><%=rds("sys_type")%> <%=rds("sys_value")%></OPTION>
											<%rds.movenext
											wend 
											rds.close 
											set rds=nothing 
											%>				
										</SELECT>
										<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(0) > (VN)越南&nbsp;
										<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(1) > (W)境外
										<input size=1 name=place type=hidden value="">
										<div><input type="checkbox" name="nc" value="C" >(N)不扣全勤<div>
									</TD>
								</TR>
								<TR>
									<TD align=right >請假日期(起)<BR><FONT class="txt8">Ngay phep (tu)</font></TD>
									<TD ><INPUT type="text" style="width:100px" NAME=HHDAT1  MAXLENGTH=10   onblur="date_change(1)"></TD>
									<TD align=right>時間<BR><FONT class="txt8">Thoi gian</font></TD>
									<TD><INPUT type="text" style="width:100px" NAME=HHTIM1 MAXLENGTH=5 ONBLUR="TIMEUP_chg()"></TD>
								</TR>	
								<TR>
									<TD align=right  >請假日期(訖)<BR><FONT class="txt8">Ngay phep (den)</font></TD>
									<TD ><INPUT type="text" style="width:100px" NAME=HHDAT2 MAXLENGTH=10   onblur="date_change(2)" ></TD>
									<TD  align=right >時間<BR><FONT class="txt8">Thoi gian</font></TD>
									<TD ><INPUT type="text" style="width:100px" NAME=HHTIM2 MAXLENGTH=5 ONBLUR="TIMEDOWN_chg()" ></TD>
								</TR> 
								<TR>
									<TD align=right>時數<BR><FONT class="txt8">Gio</font></TD>
									<TD COLSPAN=3 nowrap>
										<INPUT type="text" NAME=toth MAXLENGTH=2  STYLE="width:100px;TEXT-ALIGN:CENTER" readonly  >
										小時 =
										<INPUT type="text" NAME=totD MAXLENGTH=2  STYLE="width:100px;TEXT-ALIGN:CENTER" readonly >
										天
									</TD>
								</TR>	
								<TR>
									<TD align=right>事由<BR><FONT class="txt8">Ly do</font></TD>
									<TD COLSPAN=3><TEXTAREA rows=5 cols=60 name=memo ></TEXTAREA></TD>
								</TR>
								<tr height="50px">															
									<td align=CENTER colspan=4>
										<input type="button" name="send" value="(Y)確定Confirm" class="btn btn-sm btn-danger" onclick="GO()"  onkeydown="go()" >
										<input type="button" name="send" value="(N)取消Cancel" class="btn btn-sm btn-outline-secondary" onclick="clr()">	
										<input type="HIDDEN" size=5 name="HDcnt" value="0">
									</td>
								</tr>
							</table>
						</td>
					</tr>					
				</table>
			
</form>
  



</body>
</html>
<script language=vbscript>

function BACKMAIN()	
	open "../main.asp" , "_self"
end function   

function clr()
	open "<%=self%>.asp" , "_self"	
end function 

function typechg(a)
	if a=0 then 
		<%=self%>.place.value=""
	elseif a=1 then 
		<%=self%>.place.value="W"
	else
		<%=self%>.place.value=""	
	end if 
	
end function 

FUNCTION CALCHOUR()
	D1=<%=self%>.HHDAT1.value
	D1T=<%=self%>.HHTIM1.value 
	D2=<%=self%>.HHDAT2.value
	D2T=<%=self%>.HHTIM2.value
	if D1<>"" AND D2<>"" then 
		open "<%=self%>.back.asp?ftype=dayschg&code="& D1 &"&code1=" & D2, "Back" 
		'parent.best.cols="99%,1%"
	end if 
	if D1<>"" AND D1T<>"" AND D2<>"" AND D2T<>"" THEN 
		open "<%=self%>.back.asp?ftype=dayschg&code="& D1 &"&code1=" & D2, "Back" 
		parent.best.cols="99%,1%"
		DD1 = D1&" "&D1T
		DD2 = D2&" "&D2T
		'ALERT DD1
		'ALERT DD2
		IF ( D2T < D1T or D2<D1) THEN 
			'TOTH=FIX((DATEDIFF("N",DD1, DD2)/60))
			ALERT "起訖日期時間輸入錯誤，請假結束日期不可小於開始日期!!"
			<%=self%>.HHDAT2.value=""
			<%=self%>.HHTIM2.value=""			
			<%=self%>.HHDAT2.FOCUS()
		ELSE
			TOTH= ROUND( ROUND(DATEDIFF("N",DD1, DD2)/30,0) / 2  ,1)
			'alert TOTH
			IF TOTH>8 AND TOTH < 24 THEN 
				TOTH = 8				
			ELSEIF TOTH <= 8 AND TOTH > 4 THEN 
				TOTH = TOTH-1				
			END IF
			'TOTD=0 
			IF 	TOTH>24 THEN 
				HDcnt = cdbl(<%=self%>.HDcnt.value) 
				'alert HDcnt
				if <%=self%>.HOLIDAY_TYPE.value = "I" and <%=self%>.country.value<>"VN"  then 
					TOTH = (FIX((DATEDIFF("D",DD1, DD2)))+1)*8 
				else
					TOTH = (FIX((DATEDIFF("D",DD1, DD2)))+1)*8 - (HDcnt*8)
				end if 	
				'TOTD = (FIX((DATEDIFF("D",DD1, DD2)))) 
			END IF
			
			
			<%=self%>.totd.value = round(TOTH/8.0,1) 
		END IF	 
		'請假時數
		<%=SELF%>.TOTH.VALUE=TOTH  
		'<%=SELF%>.TOTD.VALUE=TOTD  
	ELSE
		<%=SELF%>.TOTH.VALUE=0
		'<%=SELF%>.TOTD.VALUE=0
	END IF
end FUNCTION 

 

'*******檢查日期*********************************************
FUNCTION date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.HHDAT1.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.HHDAT2.value)
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.HHDAT1.value=ANS
			CALL CALCHOUR() 
		elseif a=2 then
			Document.<%=self%>.HHDAT2.value=ANS		 
			CALL CALCHOUR()
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.HHDAT1.value=""
			Document.<%=self%>.HHDAT1.focus()
		elseif a=2 then
			Document.<%=self%>.HHDAT2.value=""
			Document.<%=self%>.HHDAT2.focus()
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
		formatDate = Year(d) & "/"  & Right("00" & Month(d), 2) & "/" & Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________  

function TIMEUP_chg() 
	IF TRIM(<%=SELF%>.HHTIM1.VALUE)<>"" THEN 
		IF ( LEFT(<%=SELF%>.HHTIM1.VALUE,2)>="24" OR RIGHT(<%=SELF%>.HHTIM1.VALUE,2)>="60" ) OR LEN(<%=SELF%>.HHTIM1.VALUE)>5  THEN 
			ALERT "時間格式輸入錯誤!!"  
			<%=SELF%>.HHTIM1.VALUE=""
			<%=SELF%>.THHTIM1.FOCUS() 
			CALL CALCHOUR()
			EXIT FUNCTION 
		ELSE
			<%=SELF%>.HHTIM1.VALUE=LEFT(<%=SELF%>.HHTIM1.VALUE,2)&":"&RIGHT(<%=SELF%>.HHTIM1.VALUE,2) 
			CALL CALCHOUR()
		END IF 		
	END IF 		 	
End function   

function TIMEDOWN_chg() 
	IF TRIM(<%=SELF%>.HHTIM2.VALUE)<>"" THEN  		 
		IF ( LEFT(<%=SELF%>.HHTIM2.VALUE,2)>="24" OR RIGHT(<%=SELF%>.HHTIM2.VALUE,2)>="60" ) OR LEN(<%=SELF%>.HHTIM2.VALUE)>5  THEN 
			ALERT "時間格式輸入錯誤!!"  
			<%=SELF%>.HHTIM2.VALUE=""
			<%=SELF%>.HHTIM2.FOCUS() 
			CALL CALCHOUR()
			EXIT FUNCTION 
		ELSE
			<%=SELF%>.HHTIM2.VALUE=LEFT(<%=SELF%>.HHTIM2.VALUE,2)&":"&RIGHT(<%=SELF%>.HHTIM2.VALUE,2)
			CALL CALCHOUR()
		END IF 		
	END IF 	 	
End function   

function GO()
	IF <%=SELF%>.EMPID.VALUE="" THEN 
		ALERT "請輸入員工編號!!" 
		<%=SELF%>.EMPID.FOCUS()
		EXIT function 
	END IF
	IF <%=SELF%>.toth.VALUE="" OR <%=SELF%>.toth.VALUE="0" THEN 
		ALERT "請輸入請假時間!!" 
		<%=SELF%>.HHDAT1.FOCUS()
		EXIT function 
	END IF 
	
	if <%=self%>.country.value<>"VN"  and (<%=self%>.radio1(0).checked=false and <%=self%>.radio1(1).checked=false ) then 
		alert "非越籍員工，請選擇休假地點為越南境內或境外!!"
		exit function 
	end if 
		
	<%=self%>.action="<%=self%>.InsertDB.asp"
	<%=self%>.submit()
END function 
	 
function jiatypechg()
	if <%=self%>.HOLIDAY_TYPE.value="I" and <%=self%>.country.value<>"VN" then 
		<%=self%>.place.value="W"
		<%=self%>.radio1(1).checked=true 
	else	
		if <%=self%>.country.value="VN" then 
			<%=self%>.place.value=""
			<%=self%>.radio1(0).checked=true 
		else
			<%=self%>.radio1(0).checked=false 
			<%=self%>.radio1(1).checked=false 
			<%=self%>.place.value=""
		end if 
	end if 
end function 	 
</script>

