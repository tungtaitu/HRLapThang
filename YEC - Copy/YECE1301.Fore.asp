<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
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
ct = request("ct")
whsno = request("whsno")

gTotalPage = 1
PageRec = 6    'number of records per page
TableRec = 20    'number of fields per record 

Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 

sql="select * from empnzjj_set where years='"& years &"' and whsno like'"&whsno&"%' and country like'"&ct&"%'   order by years, whsno,country, grade  "
		'response.write sql 
		'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 		
		pagerec= rs.RecordCount 				
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
		'whsno = rs("whsno")
	END IF 	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 
	for i = 1 to gTotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("whsno"))
				tmpRec(i, j, 2) = trim(rs("years"))
				tmpRec(i, j, 3) = trim(rs("country"))
				tmpRec(i, j, 4) = rs("grade")
				tmpRec(i, j, 5) = rs("days")
				tmpRec(i, j, 6) = rs("hs")
				tmpRec(i, j, 7) = rs("memos")				
				tmpRec(i, j, 8) = rs("aid")
				tmpRec(i, j, 9) = rs("kj")
				rs.MoveNext 
			else 
				'exit for 				
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = request("whsno")
				tmpRec(i, j, 2) = request("years")
				tmpRec(i, j, 3) = request("CT")
			end if 
			'response.write tmpRec(i, j, 0) &","&tmpRec(i, j, 2)
		next 	
		 ' if rs.EOF then 
			' rs.Close 
			' Set rs = nothing
			' exit for 
		 ' end if 			
		'Session("YECE1301") = tmpRec		 
	next 	
end if  
set rs=nothing  

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
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="100%" BORDER=0 cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr class="txt">	
									<TD nowrap  align=right height=30 >年度<br>(Nam)</TD>			
									<TD   > 
										<input type="text" style="width:100px" name="years"  maxlength=4 value="<%=years%>" onchange="gos()">
									</td>			
									<TD nowrap   align=right   >廠別<br>(Xuong)</TD>
									<TD  > 
										<select name="WHSNO"    onchange="gos()" style="width:120px">
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
									<TD   > 
										<%sql="select *from basiccode where func='country' order by sys_type" 
										set rst=conn.execute(sql)				
										%>
										<select name="ct"  onchange="gos()" style="width:120px" > 
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
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%">
								<tr height=22  bgcolor="#e4e4e4" class="txt">
									<Td align="center" width=30>STT</td>						
									<Td align="center" width=30>刪<br>除</td>	
									<Td align="center"  >年度<br>Nam</td>
									<Td align="center"  >廠別<br>Xuong</td>
									<Td align="center"  >國籍<br>Quoc tich</td>
									<Td align="center" >代碼<br>ma do</td>
									<Td align="center" >考績<br>grade</td>
									<Td align="center"  >發放<br>天數<br>so Ngay</td>			
									<Td align="center" >係數<br>he so</td>
									<Td align="center">說明<br>Ghi Chu</td>		
									
								</tr>
								<%for x = 1 to pagerec 
								if x mod 2 = 0 then 
									wkclr="#ffffff"
								else			
									wkclr="ffffff"
								end if 	
								%>
									<Tr bgcolor="<%=wkclr%>" class="txt">
										<Td align="center"><%=x%>
										<input name="aid"  value="<%=tmprec(currentpage,x,8)%>"  type="hidden">
										</td>				
										<Td >
											<input type="checkbox" name="func" onclick="del(<%=x-1%>)" >
											<input type="hidden" name="op" value="" >				
										</td>	
										<Td ><input type="text" name="nam"  value="<%=tmprec(currentpage,x,2)%>" class="readonly" readonly  style="width:100%"></td>	
										<Td ><input type="text" name="w1"  value="<%=tmprec(currentpage,x,1)%>" class="readonly" readonly  style="width:100%"></td>	
										<Td ><input type="text" name="country"  value="<%=tmprec(currentpage,x,3)%>" class="readonly" readonly  style="width:100%"></td>
										<Td><input type="text" name="grade"  value="<%=tmprec(currentpage,x,4)%>" class="readonly"  readonly  style="width:100%" style="text-align:center"></td>
										<Td><input type="text" name="Kj"  value="<%=tmprec(currentpage,x,9)%>" class="inputbox"   style="width:100%" style="width:100%;text-align:center"></td>
										<Td><input type="text" name="days"  value="<%=tmprec(currentpage,x,5)%>" class="inputbox"   style="width:100%" style="width:100%;text-align:center" onblur="dayschg(<%=x-1%>)"></td>
										<Td><input type="text" name="hs" value="<%=tmprec(currentpage,x,6)%>" class="readonly" readonly style="width:100%"  style="width:100%;text-align:center"></td>
										<Td style="width:200px"><input type="text" name="memos" value="<%=tmprec(currentpage,x,7)%>" class="inputbox" style="width:100%" ></td>				
									</tr>
								<%next%>
								<input type="hidden" name="func" value="" >				
								<input type="hidden" name="op" value="" >				
								<input type="hidden" name="aid" value="" >	
								<input type="hidden" name="w1" value="" >			
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr> 
									<TD  ALIGN=center nowrap>		
									<input type="BUTTON" name="send" value="(Y)Confirm" class="btn btn-sm btn-danger" ONCLICK="GO()">
									<input type="BUTTON" name="send" value="(N)Cancel" class="btn btn-sm btn-outline-secondary" onclick="clr()">&nbsp;
									<input type="BUTTON" name="send" value="(C)Data Copy" class="btn btn-sm btn-outline-secondary" onclick="Gocopy()">
									<input type="BUTTON" name="send" value="(S)查詢K.Tra" class="btn btn-sm btn-outline-secondary" onclick="gok()">
									</TD>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<%set conn=nothing	%>
</body>
</html>
 
<!-- #include file="../Include/func.inc" -->
<script language=vbs> 
function f()
	if <%=self%>.years.value="" then 
		<%=self%>.years.focus()
	elseif <%=self%>.whsno.value="" then 	
		<%=self%>.whsno.focus()
	elseif <%=self%>.ct.value="" then 	
		<%=self%>.ct.focus()
	else
		<%=self%>.grade(0).focus()
	end if 
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
	<%=self%>.action="<%=self%>.fore.asp"
	<%=self%>.submit()
end function 

function gok()
	wt = (window.screen.width )*0.6
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02 

	open  "yece1301.show.asp" , "_blank", "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"		 
	
end function  

function del(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="D"
	else
		<%=self%>.op(index).value=""
	end if 
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
	if <%=self%>.years.value="" then 
		alert "請輸入年度"
		<%=self%>.years.focus()
		exit function 
	end if 	
	'if <%=self%>.whsno.value="" then 
	'	alert "請輸入廠別"
	'	<%=self%>.whsno.focus()
	'	exit function 
	'end if  

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