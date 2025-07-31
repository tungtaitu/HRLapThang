<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

self="YECE1301"  
nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="1" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="1" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 	

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")


years = request("years")
y2 = request("y2")
if request("y2")="" then y2=years 
ct = request("ct")
whsno = request("whsno")

gTotalPage = 1
PageRec = 6    'number of records per page
TableRec = 50    'number of fields per record 

Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 

sql="exec proc_ShowNZJJ '"&years&"', '"& y2 &"','"&whsno&"', '"&ct&"' "
		response.write sql 
		'response.end 
 ix = 0 		
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 	
		while not rs.eof 
			ix   = ix + 1  
			rs.movenext 
		wend 
		rs.movefirst 
		pagerec= ix 
		rs.PageSize = PageRec 
		RecordInDB = ix
		TotalPage = 1
		gTotalPage = 1 
		'whsno = rs("whsno")
		cols = rs("cols")
	END IF 	
	'response.write ix 	
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 
	for i = 1 to gTotalPage 
		for j = 1 to PageRec			
			if not rs.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("whsno"))
				tmpRec(i, j, 2) = trim(rs("years"))
				tmpRec(i, j, 3) = trim(rs("country"))							
				for  k = 1 to cols  
					tbn1 = chr(k+64)&"1" 
					tbn2 = chr(k+64)&"1days" 
					tbn3 = chr(k+64)&"1hs"  	 
					tbn4 = chr(k+64)&"1kj"
					tmpRec(i, j, 4+(k-1)*4) = rs(tbn1)
					tmpRec(i, j, 5+(k-1)*4) = rs(tbn2)
					tmpRec(i, j, 6+(k-1)*4) = rs(tbn3)			
					tmpRec(i, j, 7+(k-1)*4) = rs(tbn4)		
					'response.write tmpRec(i, j, 4)&","& tmpRec(i, j, 5)&"," & tmpRec(i, j, 6) &"<BR>"
				next  
				'response.write k &"<BR>"  
				' tmpRec(i, j, 7) = rs("memos")				
				' tmpRec(i, j, 8) = rs("aid")
				
				rs.MoveNext 
				'k=1
			else 								
				exit for 	
			end if 
			'response.write tmpRec(i, j, 0) &","&tmpRec(i, j, 2)
		next 	
		 if rs.EOF then 
			rs.Close 
			Set rs = nothing
			exit for 
		 end if 			
		'Session("YECE1301") = tmpRec		 
	next 	
end if  
set rs=nothing  
'response.end 

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">

</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()"  >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=SESSION("PGNAME")%></TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=500  ><tr><td >
	<table width=500 align=center border=0 cellspacing="1" cellpadding="1"  class="txt8" >
		<tr>	
			<TD nowrap  align=right height=30 >年度<br>(Nam)</TD>			
			<TD    > 
				<input name="years" class="inputbox" size=5  maxlength=4 value="<%=years%>"  >~
				<input name="Y2" class="inputbox" size=5  maxlength=4 value="<%=y2%>"  >
				<input  type="button" name="btn" value="(查詢)"  class="button" onclick="gos()" onkeydown="gos()" >				
			</td>		 
			<TD nowrap   align=right   >廠別<br>(Xuong)</TD>
			<TD  > 
				<select name="WHSNO"  class="txt8"   onchange="gos()" >
					<option value="">----</option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if whsno=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD> 		

			<TD nowrap  align=right height=30 >國籍<br>(Quoc tich)</TD>			
			<TD  nowrap > 
				<%sql="select *from basiccode where func='country' order by sys_type" 
				set rst=conn.execute(sql)				
				%>
				<select name="ct" class="txt8"  onchange="gos()"  > 
					<option value="">----</option>
					<%while not rst.eof%>
						<option value="<%=rst("sys_type")%>"  <%if ct=rst("sys_type") then%>selected<%end if%> ><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
					<%rst.movenext
					wend
					set rst=nothing 
					%>	
				</select> 
				
			</td>			 			
		</TR>					
	</table>	
	 
	<table   align=center border=0 cellspacing="1" cellpadding="2"  class="txt" >
		<tr height=22  bgcolor="#e4e4e4"  class="txt8">
			<Td align="center" width=30  nowrap  >STT</td>
			<Td align="center" width=50 nowrap  >年度<br>Nam</td>
			<Td align="center" width=50 nowrap >廠別<br>Xuong</td>
			<Td align="center" width=50 nowrap  >國籍<br>Quoc tich</td>
			<%for zz = 1 to cols %>
				<Td align="center" width=50 nowrap   >考績<br>天數</td>
			<%next%>	
		</tr> 
		
		<%for x = 1 to pagerec 
		if x mod 2 = 0 then 
			wkclr="#ffcccc"
		else			
			wkclr="#ffcccc"
		end if 	
		%>
			<%if f_whsno<>"" and tmprec(currentpage,x,1)<>f_whsno then%>
			<tr>
				<td colspan=<%=4+cols%>><hr size=0	style='border: 1px dotted #999999;' align=left ></td>
			</tr>
			<%end if%>
			<Tr bgcolor="<%=wkclr%>">
				<Td align="center" width=30 align="center"><%=x%></td>				
				<Td  align="center"><%=tmprec(currentpage,x,2)%></td>				 
				<Td  align="center"><%=tmprec(currentpage,x,1)%></td>	
				<Td   align="center"><%=tmprec(currentpage,x,3)%></td>			
				<%for zz = 1 to cols %>
				<td align="center"><%=tmprec(currentpage,x,7+(zz-1)*4)%> <br>
					<font color="<%if tmprec(currentpage,x,5+(zz-1)*4)="0" then %>#ffcccc<%end if%>" class="txt12">
					<B><%=tmprec(currentpage,x,5+(zz-1)*4)%><b></font>					
				</td>
				<%next%>
			</tr>		
			<%f_whsno=tmprec(currentpage,x,1)%>			
		<%next%>
	</table>
	<br>
	<Table width=550>
		<tr> 
			<TD  ALIGN=center nowrap>		
			<input type="BUTTON" name="send" value="(X)關閉視窗Close" class=button ONCLICK="window.close()">			
			</TD>
		</tr>
	</table>	
<hr size=0	style='border: 1px dotted #999999;' align=left >	

</td></tr></table>
<%set conn=nothing	%>
</body>
</html>
 
<!-- #include file="../Include/func.inc" -->
<script language=vbs> 
function f()
	'if <%=self%>.years.value="" then 
		<%=self%>.years.focus()	
	'end if 
end function 

function gos()
	' pg = <%=self%>.pagerec.value
	' if <%=self%>.years.value<>"" then 
		' for x = 1 to 6 
			' if trim(<%=self%>.nam(x-1).value)="" then 
				' <%=self%>.nam(x-1).value = trim(<%=self%>.years.value)
			' end if 	
		' next 
	' end if 
	<%=self%>.totalpage.value="0"
	<%=self%>.action="<%=self%>.show.asp"
	<%=self%>.submit()
end function 

function Gocopy()
	open "<%=self%>B.fore.asp" , "_self"
end function 

function clr()
	open "<%=SELF%>.asp" , "_self"
end function 

function dayschg(index)
	if <%=self%>.days(index).value<>"" then 
		if isnumeric(<%=self%>.days(index).value)=false then 
			alert "請輸入數字,xin danh lai [so] !!"
			<%=self%>.days(index).value=""
			<%=self%>.days(index).focus()
			exit function 
		else
			<%=self%>.hs(index).value = round(cdbl(<%=self%>.days(index).value)/30+0.001,2)
		end if 
	end if 	
end function 

function tot_Mtaxchg(index,a)
	if a=1 then 
		if <%=self%>.person_qty(index).value<>"" then 
			if isnumeric(<%=self%>.person_qty(index).value)=false then 
				alert "請輸入數字,xin danh lai [so] !!"
				<%=self%>.person_qty(index).value=""
				<%=self%>.person_qty(index).focus()
				exit function 
			end if 
		end if 	
	elseif a=2 then 
		if <%=self%>.ut_mtax(index).value<>"" then 
			if isnumeric(<%=self%>.ut_mtax(index).value)=false then 
				alert "請輸入數字,xin danh lai [so] !!"
				<%=self%>.ut_mtax(index).value=""
				<%=self%>.ut_mtax(index).focus()
				exit function 
			else	
				<%=self%>.ut_mtax(index).value=formatnumber(<%=self%>.ut_mtax(index).value,0)
			end if 
		end if 	
	end if 
	
	if trim(<%=self%>.ut_mtax(index).value)<>"" and trim(<%=self%>.person_qty(index).value)<>"" then 
		<%=self%>.tot_Mtax(index).value=formatnumber( cdbl(<%=self%>.person_qty(index).value)*cdbl(<%=self%>.ut_mtax(index).value) , 0)
	end if 
end function 

function empidchg(index)
	code1=UCase(trim(<%=self%>.empid(index).value))
	if <%=self%>.empid(index).value<>"" then 		
		open "<%=self%>.back.asp?func=chkemp&code1="& code1 &"&index="& index  , "Back"
		'parent.best.cols="50%,50%"
	end if 
end function 
function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go() 
	if <%=self%>.whsno.value="" then 
		alert "請輸入廠別"
		<%=self%>.whsno.focus()
		exit function 
	end if 
	<%=self%>.action="<%=SELF%>.upd.asp"
 	<%=self%>.submit() 
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF 
END FUNCTION
</script> 