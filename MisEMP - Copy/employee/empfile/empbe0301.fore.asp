<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<!-- #include file="../../Include/SIDEINFO.inc" -->
<%
self="empbe0301"  

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date()-1)&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 

Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   

s_empid=request("s_empid")
s_dat1=request("s_dat1")
s_dat2=request("s_dat2")
s_country=request("s_country") 


'response.write s_country  


gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 15    'number of fields per record     

sql=" select  b.groupid, b.gstr, b.empnam_cn,  convert(char(10), a.mdtm, 111) as keyindat , "&_
	"convert(char(10), sdat, 111) as sdate, convert(char(10), edat, 111) as edate , a.* from "&_
	"( select * from empvisadata  ) a  "&_
	"join ( select  * from view_empfile ) b on b.empid  = a.empid  "&_
	"where a.empid like '%"& s_empid &"%' and  b.country like '"& s_country &"%' " 
if s_dat1<>"" then 
	sql=sql & "and convert(char(10), a.edat, 111) between '"& s_dat1 &"' and '"& s_dat2 &"' " 
end if 	

sql=sql & "order by	a.empid , a.edat "  

'response.write sql 
'response.end 

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1 	 	  		
	rs.Open SQL, conn, 3, 3  
	IF NOT RS.EOF THEN
		PageRec = rs.RecordCount 
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
				tmpRec(i, j, 3) = trim(rs("groupid"))
				tmpRec(i, j, 4) = trim(rs("gstr"))
				tmpRec(i, j, 5) = trim(rs("keyindat"))
				tmpRec(i, j, 6) = trim(rs("visano"))
				tmpRec(i, j, 7) = trim(rs("sdate"))
				tmpRec(i, j, 8) = trim(rs("edate"))
				tmpRec(i, j, 9) = trim(rs("visaamt"))
				tmpRec(i, j, 10) = trim(rs("memo"))
				tmpRec(i, j, 11) = trim(rs("aid"))
				tmpRec(i, j, 12) = trim(rs("country"))
				rs.movenext
			else
				exit for			
			end if 
	
			if rs.EOF then
				rs.Close
				Set rs = nothing
				exit for
			 end if
		next
	next 
end if 		
		 	
  


%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.s_empid.focus()	
	<%=self%>.s_empid.select()	
	'<%=self%>.country.SELECT()
end function    

function resch()
	<%=self%>.totalpage.value="0" 
	<%=self%>.action="<%=self%>.fore.asp"
	<%=self%>.submit()
end function 
-->
</SCRIPT>   
</head> 
<body  topmargin="40" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=pagerec VALUE="<%=pagerec%>"> 
<table width="460" border="0" cellspacing="0" cellpadding="0">
	 
	<tr><td colspan=3><hr size=0	style='border: 1px dotted #999999;' align=left width=580></td></tr>
	<tr height=40 >
		<td width=150 align=center valign=middle>
			<img border="0" src="../../picture/icon02.gif" align="absmiddle"> 
			<a href="empbe03.asp" target="_parent">簽證資料新增3</a>
		</td>
		<td width=180 align=center>
			
			<font color="Brown"><b>簽證資料異動查詢</b></font>
		</td>
		<td></td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=580>
<fieldset style="margin:0;padding:0;width=580"><legend><font class=txt9>資料查詢</font></legend>
	<table width=570 class=txt9 border="0" cellspacing="0" cellpadding="0" >	
		<tr height=35>
			<td width=50 align=right>工號:</td>
		 	<td width=60>
		 		<input name=S_empid class=inputbox size=8 value="<%=S_empid%>">	 		
		 	</td>
			<td width=80 align=right>簽證有效期:</td>
		 	<td width=180>
		 		<input name=S_dat1 class=inputbox value="<%=s_dat1%>" size=11 onblur="date_change(1)">~
		 		<input name=S_dat2 class=inputbox value="<%=s_dat2%>" size=11 onblur="date_change(2)">	 		
		 	</td>
		 	
		 	<td width=40 align=right>國籍:</td>
		 	<td>
		 		<select name=s_country class=inputbox onchange="resch()">
		 			<option value="">全部</option>
		 			<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type <>'VN' ORDER BY SYS_TYPE "
						SET RST = CONN.EXECUTE(SQL)
						WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=s_country THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
					<%
						RST.MOVENEXT
						WEND
						set rst=nothing
					%>
		 		</select>
		 	</td>
		 	<td align=cneter>
		 		<input type=button name=send value="查詢" class=button onclick="resch()"  onkeydown="resch()" >
		 	</td>
		</tr>	
	</table>	
</fieldset>		
<table width=500  ><tr><td >
	<table width=480 class=txt9>
		<tr bgcolor="#DCDCDC" height=25>
			<td align=center>DEL</td>
			<td align=center>工號</td>
			<td align=center>姓名</td>
			<td align=center>簽証號碼</td>
			<td align=center>有效期(起)</td>
			<td align=center>有效期(迄)</td>
			<td align=center>費用(VND)</td>
			<td align=center>備註</td>
		</tr>
		<% for CurrentRow = 1 to PageRec %>
		<tr>
			<td>
				<%if  tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>
					<input type=checkbox name=btn value="DEL" onclick="delchg(<%=currentRow-1%>)" >			
				<%else%>	
					<input type=hidden name="btn">
				<%end if %>
			</td>
			<td>
				<input name=empid size=6 class=readonly readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" >
				<input type=hidden name=aid value="<%=tmpRec(CurrentPage, CurrentRow, 11)%>">
			</td>	
			<td>
				<input name=empname size=15 class=readonly readonly value="<%=tmpRec(CurrentPage, CurrentRow, 2)%>" >
				<input type=hidden name=country  value="<%=tmpRec(CurrentPage, CurrentRow, 12)%>" >
			</td>	
			<td>
				<input name=visano size=10 class=inputbox  onblur="visanochg(<%=CurrentRow-1%>)" maxlength=8 value="<%=tmpRec(CurrentPage, CurrentRow, 6)%>" >
			</td>	
			<td>
				<input name=dat1 size=11 class=inputbox onblur="dat1chg(<%=CurrentRow-1%>)" value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>">
			</td>	
			<td>
				<input name=dat2 size=11 class=inputbox onblur="dat2chg(<%=CurrentRow-1%>)" value="<%=tmpRec(CurrentPage, CurrentRow, 8)%>">
			</td>	
			<td>
				<input name=visaAmt size=10 class=inputbox onblur="amtchg(<%=CurrentRow-1%>)"  style='text-align:right' value="<%=tmpRec(CurrentPage, CurrentRow, 9)%>" >
			</td>	
			<td>
				<input name=memo size=15 class=inputbox onblur="visanochg(<%=CurrentRow-1%>)" value="<%=tmpRec(CurrentPage, CurrentRow, 10)%>">
			</td>	
		</tr>
		<%next%>
	</table>	
	<input type=hidden name="btn">
	<input type=hidden name="empid">
	<input type=hidden name="aid">
	<input type=hidden name="empname">
	<input type=hidden name="country">
	<input type=hidden name="visano">
	<input type=hidden name="dat1">
	<input type=hidden name="dat2">
	<input type=hidden name="visaamt">
	<input type=hidden name="memo">
	<table width=450 align=center>
		<tr class=txt9>
			<td align=center>---資料共 <%=RecordInDB%> 筆---</td>
		</tr>		
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="(Y)Confirm 修 改" onclick="go()" onkeydown="go()">
				<input type=button  name=btm class=button value="(N)Cancel"  onclick="JavaScript:window.location.reload()">
			</td>
		</tr>	
	</table>	

</td></tr></table> 

</body>
</html>


<script language=vbs> 
function getempdata(index) 
	ncols="visano"
	open "Getempdata.asp?pself="& "<%=self%>" &"&index=" & index &"&ncols="& ncols , "Back" 
	parent.best.cols="50%,50%"
end function   

function delchg(index)
	if confirm("確定要刪除這筆資料Delete(Cancel) This Record?",64) then 
		aid = <%=self%>.aid(index).value 
		
		<%=self%>.action ="<%=self%>.deldb.asp?code1="& aid 
		<%=self%>.submit()
	end if 	
end function  


function chkempid(index)	
	if <%=self%>.empid(index).value<>"" then 
		code1=Ucase(trim(<%=self%>.empid(index).value))
		open "<%=self%>.back.asp?func=chkempid&index=" & index &"&code1=" & code1 , "Back" 
		'parent.best.cols="70%,30%"
	end if 
end  function 

function visanochg(index)	
	<%=self%>.visano(index).value=Ucase(<%=self%>.visano(index).value)
	<%=self%>.memo(index).value=Ucase(<%=self%>.memo(index).value)
end  function  

function amtchg(index)	
	if <%=self%>.visaAmt(index).value<>"" then 	
		if isnumeric(<%=self%>.visaAmt(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.visaAmt(index).value="0"
			<%=self%>.visaAmt(index).focus()
			<%=self%>.visaAmt(index).select()
		end  if 
	end if 		
end  function   


function go() 
	empstr=""
	recd=<%=self%>.RecordInDB.value 
	for x = 1 to recd+1 
		if <%=self%>.empid(x-1).value<>"" then 
			if <%=self%>.visaNo(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 簽證號碼!!"
				<%=self%>.visaNo(x-1).focus()
				exit function
			elseif 	<%=self%>.dat1(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 有效期(起)!!"
				<%=self%>.dat1(x-1).focus()
				exit function
			elseif 	<%=self%>.dat2(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 有效期(迄)!!"
				<%=self%>.dat2(x-1).focus()
				exit function
			'elseif 	<%=self%>.visaamt(x-1).value=""  or  <%=self%>.visaamt(x-1).value="0" then 
			'	alert "請輸入 "&<%=self%>.empid(x-1).value&" 簽證費用!!"
			'	<%=self%>.visaamt(x-1).focus()	
			'	<%=self%>.visaamt(x-1).select()
			'	exit function
			end if 
		end  if
		empstr = empstr & Ucase(<%=self%>.empid(x-1).value)
	next 
	if len(empstr)=0 then 
		alert "請輸入資料!!"
		<%=self%>.empid(0).focus()
		exit function 
	else	
	 	<%=self%>.action="<%=self%>.Upd.asp"
	 	<%=self%>.submit() 
	 end  if 	
end function   
	

'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.s_dat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.S_dat2.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.S_dat1.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.S_dat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.S_dat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.S_dat2.value=ANS
		elseif a=3 then
			Document.<%=self%>.S_dat1.value=ANS
		elseif a=4 then
			Document.<%=self%>.S_dat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.S_dat1.value=""
			Document.<%=self%>.S_dat1.focus()
		elseif a=2 then
			Document.<%=self%>.S_dat2.value=""
			Document.<%=self%>.S_dat2.focus()
		elseif a=3 then
			Document.<%=self%>.S_dat1.value=""
			Document.<%=self%>.S_dat1.focus()
		elseif a=4 then
			Document.<%=self%>.S_dat2.value=""
			Document.<%=self%>.S_dat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION   


'*******檢查日期*********************************************
FUNCTION dat1chg(index)	

	INcardat = Trim(<%=self%>.dat1(index).value)  		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN		
		Document.<%=self%>.dat1(index).value=ANS					
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 		
		Document.<%=self%>.dat1(index).value=""
		Document.<%=self%>.dat1(index).focus() 		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION 

FUNCTION dat2chg(index)	
	INcardat = Trim(<%=self%>.dat2(index).value)  		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN		
		Document.<%=self%>.dat2(index).value=ANS					
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 		
		Document.<%=self%>.dat2(index).value=""
		Document.<%=self%>.dat2(index).focus() 		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF    
END FUNCTION   


</script> 