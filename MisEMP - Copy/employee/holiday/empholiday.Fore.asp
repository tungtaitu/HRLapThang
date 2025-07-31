<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
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
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
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
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD >
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	員工請假作業 </TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<TABLE WIDTH=460 BORDER=0>   
	<TR height=25 >
		<TD nowrap align=right WIDTH=150>員工編號<BR><FONT class="txt8">so the</font></TD>
		<TD COLSPAN=3 >
			<INPUT NAME=EMPID SIZE=10 CLASS=INPUTBOX  onchange="empidchg()"> 
			<INPUT NAME="NVx" SIZE=10  CLASS="READONLY8" READONLY style="height:22" >
			<INPUT NAME="ntv" SIZE=15  CLASS="READONLY8" READONLY style="height:22" >
			<INPUT NAME="country" SIZE=5  CLASS="READONLY8" READONLY style="height:22" type="hidden" >
		</TD> 
	</TR>	
	<TR>
		<TD></TD>
		<TD COLSPAN=3 ><INPUT NAME=EMPNAMEVN SIZE=40  CLASS="READONLY8"   READONLY style="height:22" > </TD>
	</TR>	
	<TR><TD HEIGHT=20></TD></TR>
	<TR>
		<TD align=right >尚有年假<BR><FONT class="txt8">Giao PN</font></TD>
		<TD COLSPAN=3 > 
			<input name="ynjh" size=5 class=readonly readonly value="" style="text-align:center">
			<input name="ynj" size=35 class=readonly readonly value="">
		</TD>
	</TR>	
	<TR>
		<TD align=right >請假類別<BR><FONT class="txt8">loai phep</font></TD>
		<TD COLSPAN=3>
			<SELECT NAME=HOLIDAY_TYPE CLASS=TXT11 onchange="jiatypechg()"> 
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
			<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(0) > 越南&nbsp;
			<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(1) > 境外
			<input size=1 name=place type=hidden value="">
		</TD>
	</TR>
	<TR>
		<TD align=right   WIDTH=150>請假日期(起)<BR><FONT class="txt8">nagy phep (tu)</font></TD>
		<TD WIDTH=80 ><INPUT NAME=HHDAT1 CLASS=INPUTBOX SIZE=12 MAXLENGTH=10   onblur="date_change(1)"></TD>
		<TD align=right WIDTH=60>時間<BR><FONT class="txt8">thoi giam</font></TD>
		<TD WIDTH=170><INPUT NAME=HHTIM1 CLASS=INPUTBOX SIZE=8 MAXLENGTH=5 ONBLUR="TIMEUP_chg()"></TD>
	</TR>	
	<TR>
		<TD align=right WIDTH=150 >請假日期(訖)<BR><FONT class="txt8">nagy phep (den)</font></TD>
		<TD WIDTH=80 ><INPUT NAME=HHDAT2 CLASS=INPUTBOX SIZE=12 MAXLENGTH=10   onblur="date_change(2)" ></TD>
		<TD WIDTH=60 align=right >時間<BR><FONT class="txt8">thoi giam</font></TD>
		<TD WIDTH=170><INPUT NAME=HHTIM2 CLASS=INPUTBOX SIZE=8 MAXLENGTH=5 ONBLUR="TIMEDOWN_chg()" ></TD>
	</TR> 
	<TR>
		<TD align=right VALIGN=TOP>時數<BR><FONT class="txt8">Giô</font></TD>
		<TD COLSPAN=3><INPUT NAME=toth CLASS=INPUTBOX SIZE=5 MAXLENGTH=2  STYLE="TEXT-ALIGN:CENTER" readonly  >小時 = 
		<INPUT NAME=totD CLASS=INPUTBOX SIZE=5 MAXLENGTH=2  STYLE="TEXT-ALIGN:CENTER" readonly >天
		<!--INPUT NAME=totD CLASS=READONLY READONLY  SIZE=5 MAXLENGTH=2  STYLE="TEXT-ALIGN:CENTER" -->
		</TD>
	</TR>	
	<TR>
		<TD align=right VALIGN=TOP>事由<BR><FONT class="txt8">ly do</font></TD>
		<TD COLSPAN=3><TEXTAREA rows=5 cols=60 name=memo class="txt8"></TEXTAREA></TD>
	</TR>
</table>	

<BR> 
<TABLE border=0 width=460 class=font9 >
<tr>
    <!--td align="CENTER" height=40 width=70%>
    
	<% If CurrentPage > 1 Then %>
		<input type="submit" name="send" value="FIRST" class=button>
		<input type="submit" name="send" value="BACK" class=button>
	<% Else %>
		<input type="submit" name="send" value="FIRST" disabled class=button>
		<input type="submit" name="send" value="BACK" disabled class=button>
	<% End If %>		
	<% If cint(CurrentPage) < cint(TotalPage) Then %>
		<input type="submit" name="send" value="NEXT" class=button>
		<input type="submit" name="send" value="END" class=button>
	<% Else %>      
		<input type="submit" name="send" value="NEXT" disabled class=button>
		<input type="submit" name="send" value="END" disabled class=button>	
	<% End If %>　
	PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
	</td-->
	<td align=CENTER>
		<input type="button" name="send" value="(Y)確定Confirm" class=button onclick="GO()"  onkeydown="go()" >
		<input type="button" name="send" value="(N)取消Cancel" class=button onclick="clr()">	
		<input type="HIDDEN" size=5 name="HDcnt" value="0">
	</td>		
</TR>
</TABLE> 
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
		'parent.best.cols="70%,30%"
	end if 
	if D1<>"" AND D1T<>"" AND D2<>"" AND D2T<>"" THEN 
		open "<%=self%>.back.asp?ftype=dayschg&code="& D1 &"&code1=" & D2, "Back" 
		'parent.best.cols="70%,30%"
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
			IF TOTH>8 AND TOTH < 24 THEN 
				TOTH = 8				
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

