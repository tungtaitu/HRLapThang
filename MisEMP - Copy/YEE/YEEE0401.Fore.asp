<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%response.buffer=true%>
<%
Set conn = GetSQLServerConnection()
self="yeee0401" 

if  instr(conn,"168")>0 then 
	w1="LA"
elseif  instr(conn,"169")>0 then 
	w1="DN"	
elseif  instr(conn,"47")>0 then 
	w1="BC"	
end if 	  

sortby = request("sortby") 
if sortby=""  then 
	sortby="B"
	sort_str = "a.empid, a.s_dat "	
elseif sortby="A" then 
	sort_str = "a.s_dat , a.empid"
elseif sortby="B" then 
	sort_str = "a.empid, a.s_dat "	
elseif sortby="C" then 	
	sort_str = "a.jb, a.empid, a.s_dat "	
end if 	

dd1=request("dd1")
dd2=request("dd2")
sothe=ucase(triM(request("sothe")))
ct = request("ct")
otherQ = trim(request("otherQ")) 
sts2 = trim(request("sts2"))  
whsno = trim(request("whsno"))  

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 20    'number of fields per record  

nowym=year(date())&right("00"&month(date()),2)

sqlx="select * from vyfyexrt where yyyymm='"& nowym &"' and code='USD' "
set rsx=conn.execute(sqlx)
if not rsx.eof then 
	rate = rsx("exrt")
else
	rate = 1 
end if  
set rsx=nothing 
sql="select  * from empphi  order by years desc , country "
 
Set rs = Server.CreateObject("ADODB.Recordset")

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec+5
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("years"))		
				tmpRec(i, j, 2) = trim(rs("country"))		
				tmpRec(i, j, 3) = trim(rs("phi_vnd"))		
				tmpRec(i, j, 4) = trim(rs("phi_usd"))		
				tmpRec(i, j, 5) = trim(rs("aid"))		
				tmpRec(i, j, 6) = trim(rs("Btax"))		
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
	Session("yeee0401B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("yeee0401B")

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

nowdate = year(date())&"/"&right("00"&month(date()),2)&"/"&right("00"&day(date()),2)

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
 
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=session("pgname")%></TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>

<table width=500  ><tr><td align="center" >
	<table width=450 BORDER=0 cellspacing="1" cellpadding="1" class=txt bgcolor=black>
	<tr bgcolor=#ffffff height=35>
		<Td align=center bgcolor="#ffffff" width=150 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand" ><a href="yeee04.fore.asp" target="_self">銷假作業</a></td>
		<Td align=center bgcolor="#ffcccc"  width=150 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffcccc'"   style="cursor:hand"><a href="yeee0401.asp" target="_self">差旅費設定</a></td>
		<Td align=center bgcolor="#ffffff"  width=150 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"></td>		
	</tr>	
	</table>	
	 
	<table width=500  BORDER=0 cellspacing="2" cellpadding="2" class=txt8 bgcolor="#ffffff">
		<tr>
			<td width=80 align="right">本月匯率</td>
			<td><input name="rate" class="inputbox8" size=8 value="<%=rate%>"></td>
	</table>
	<table  BORDER=0 cellspacing="1" cellpadding="2" class=txt8 bgcolor="#ffffff">
		<tr height=25 bgcolor="#e4e4e4">
			<Td width=30 nowrap align="center">STT</td>
			<Td width=50 nowrap align="center" >年度</td>
			<Td width=60 nowrap align="center" >國籍</td>
			<Td width=100 nowrap align="center" >差旅費(VND)</td>
			<Td width=50 nowrap align="center" >稅額<br>%</td>
			<Td width=100 nowrap  align="center">差旅費(USD)</td>						
		</tr>
		<%response.flush%>
		<%for x = 1 to pagerec %>
			<tr bgcolor="#ffffff">
				<td><%=x%></td>	
				<td><input name="years" size=5 class="inputbox8" value="<%=tmprec(currentpage,x,1)%>" ></td>	
				<td>
					<select name="country" class="txt8" style="width:80"  >
						<option value="">---</option>
						<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type<>'VN' ORDER BY SYS_type desc  "
						SET RST = CONN.EXECUTE(SQL)
						WHILE NOT RST.EOF  
						%>
						<option value="<%=RST("SYS_TYPE")%>" <%if tmprec(currentpage,x,2)=rst("sys_type") then%> selected<%end if%>><%=RST("SYS_TYPE")%>  - <%=RST("SYS_VALUE")%></option>				 
						<%
						RST.MOVENEXT
						WEND 
						%>
					</SELECT>
					<%SET RST=NOTHING %>
					</select>	
					</td>	
				 <td><input name="phi_vnd" size=15 class="inputbox8r" onblur="phivchg(<%=x-1%>)" value="<%=formatnumber(tmprec(currentpage,x,3),0)%>"></td>	
				 <td><input name="btax" size=5 class="inputbox8r"  value="<%=tmprec(currentpage,x,6)%>" onblur="btaxchg(<%=x-1%>)"></td>	
				 <td><input name="phi_usd" size=15 class="inputbox8r" value="<%=formatnumber(tmprec(currentpage,x,4),0)%>"></td>	
			</tr>
		<%next%>
	</table>
 
	<TABLE WIDTH=500 class=txt8>
			<tr ALIGN=center>
		    <td align="left" height=40  >
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
			<% End If %>&nbsp;
			PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
			</td>
			<td>
				<input type=button name=btn value="(Y)Confirm"  class=button onclick="go()">				
			</td>
			</TR>
	</TABLE>	
		
</td></tr></table>

</body>
</html>
<!-- #include file="../Include/func.inc" -->

<script language=vbs>
function f()
	'<%=self%>.yymm.focus()
end function  

function btaxchg(index)
	if trim(<%=self%>.btax(index).value)<>"" then 
		if isnumeric(trim(<%=self%>.btax(index).value)) = false then 
			alert "請輸入數字,"
			<%=self%>.btax(index).value="0"
			<%=self%>.btax(index).focus()
			exit function 
		else
			if  trim(<%=self%>.phi_vnd(index).value)<>"" and trim(<%=self%>.rate.value)<>"" then 
				btax = cdbl(<%=self%>.btax(index).value)
				phi_vnd = cdbl(<%=self%>.phi_vnd(index).value) 
				rate = cdbl(<%=self%>.rate.value) 
				<%=self%>.phi_usd(index).value=formatnumber( formatnumber(phi_vnd/rate ,0) * (1+(btax/100)) , 0) 
				datachg(index) 
			end if 
		end if 
	end if 	
end function 

function phivchg(index)
	if trim(<%=self%>.phi_vnd(index).value)<>"" then 
		if isnumeric(<%=self%>.phi_vnd(index).value)=false then 
			alert "需輸入數字xin danh lai so !!"
			<%=self%>.phi_vnd(index).value=0
			<%=self%>.phi_vnd(index).select()
			exit function 
		elseif cdbl(<%=self%>.phi_vnd(index).value)<0 or instr(<%=self%>.phi_vnd(index).value,".")>0  then 	
			alert "需輸入整數,不可<0 ( xin danh lai so , ko duoc <0 ) !!"
			<%=self%>.phi_vnd(index).value=0
			<%=self%>.phi_vnd(index).select()
			exit function 
		else
			<%=self%>.phi_vnd(index).value=formatnumber(<%=self%>.phi_vnd(index).value,0) 
			if <%=self%>.rate.value<>"" then 
				if isnumeric(<%=self%>.rate.value) then 
					<%=self%>.phi_usd(index).value=formatnumber( cdbl(<%=self%>.phi_vnd(index).value)/cdbl(<%=self%>.rate.value),0) 
					datachg(index)
				end if 
			end if 	
		end if 
	end if 	
end function 
 


function datachg(index)
	str1 = <%=self%>.years(index).value  
	str2 = <%=self%>.country(index).value 
	str3 = <%=self%>.phi_vnd(index).value 
	str4 = <%=self%>.phi_usd(index).value 
	str5 = <%=self%>.btax(index).value 
	 	
	open "<%=self%>.back.asp?func=datachg&code1="& str1 & "&code2="& str2 &"&code3="& str3 &"&code4="&str4 &"&code5="& str5  &"&index="& index  &"&currentpage="& <%=currentpage%>    , "Back"  
	'parent.best.cols="50%,50%" 
end function 

function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function 

function showholiday(index)
	wt = (window.screen.width )*0.8
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.05
	lt = (window.screen.availHeight)*0.1	
	empid = <%=self%>.empid(index).value
	dat1 = <%=self%>.dat1(index).value
	dat2 = <%=self%>.dat2(index).value

	open "<%=self%>.showdata.asp?empid1="& empid &"&dat1="& dat1 &"&dat2="& dat2     , "balnkN"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"
end function 

function go()	
 	<%=self%>.action="<%=self%>.upd.asp"
 	<%=self%>.submit()
end function

function goexcel()
	if <%=self%>.yymm.value="" then 
		alert "請輸入[統計年度]!!"
		<%=self%>.yymm.focus()
		exit function 
	end if 	
	'open "<%=self%>.toexcel.asp" , "Back" 
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	'parent.best.cols="50%,50%"
end function  

'*******檢查日期*********************************************
FUNCTION date_change(index,a)

if a=1 then
	INcardat = Trim(<%=self%>.dat1(index).value)
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2(index).value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1(index).value=ANS
			datachg(index)
		elseif a=2 then
			Document.<%=self%>.dat2(index).value=ANS
			datachg(index)
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.dat1(index).value=""
			Document.<%=self%>.dat1(index).focus()
		elseif a=2 then
			Document.<%=self%>.dat2(index).value=""
			Document.<%=self%>.dat2(index).focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION  

FUNCTION ddDatechg(a)

if a=1 then
	INcardat = Trim(<%=self%>.dd1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.dd2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dd1.value=ANS
		elseif a=2 then
			Document.<%=self%>.dd2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.dd1.value=""
			Document.<%=self%>.dd1.focus()
		elseif a=2 then
			Document.<%=self%>.dd2.value=""
			Document.<%=self%>.dd2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION  

function gos()
	<%=self%>.TotalPage.value = "" 
	<%=self%>.action="<%=self%>.Fore.asp"
	<%=self%>.submit()
end function 
</script> 
<%response.end%>
