<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->

<%
Set conn = GetSQLServerConnection()
self="yeee0401P" 

if  instr(conn,"168")>0 then 
	w1="LA"
	w2 = "越南"	
elseif  instr(conn,"169")>0 then 
	w1="DN"	
	w2 = "同奈"	
elseif  instr(conn,"47")>0 then 
	w1="BC"	
	w2 = "越南"	
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

xid = request("xid")
empid=request("empid")
  

nowym=year(date())&right("00"&month(date()),2)

sqlx="select * from vyfyexrt where yyyymm='"& nowym &"' and code='USD' "
set rsx=conn.execute(sqlx)
if not rsx.eof then 
	rate = rsx("exrt")
else
	rate = 1 
end if  
set rsx=nothing  


 
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
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"    >
<form name="<%=self%>" method="post" >
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<table  BORDER=0 cellspacing="0" cellpadding="0" class=txt12 width=580 > 	 
	<tr>
		<td valign="middle" width=80 align="right">
			<img border="0" src="../picture/yfylogo_a2.gif" align="absmiddle" width=70 height=58>
		</td>
		<Td align="center">
		永豐餘紙業(  <%=W2%>  )有限公司( ___<%=w1%>___ ) 廠 <br>
		出差申請及旅費報支單<br>
		PHIẾU CÔNG TÁC VÀ BÁO CHI PHÍ CẦU ĐƯỜNG
		</td>
	</tr>
</table>	
<table  BORDER=0 cellspacing="0" cellpadding="1" class=txt12 width=600 > 	 
	<tr>
		<td nowrap width=80 align="center" valign="bottom">	<span style="font-size:22pt;">□</span> 國內 	</td>		
		<td nowrap width=80 align="center" valign="bottom">	<span style="font-size:22pt;">□</span> 國外 	</td>		
		<td >	</td>		
		<td nowrap width=80 align="right" valign="bottom">	申請日期 	</td>		
		<td nowrap width=80 align="right" valign="bottom">	年 	</td>		
		<td nowrap width=60 align="right" valign="bottom">	月 	</td>		
		<td nowrap width=60 align="right" valign="bottom">	日 	</td>		
	</tr>
	<tr class="txt">
		<td align="center" valign="Top">Trong Nước</td>		
		<td align="center" valign="Top">Nước Ngoài </td>		
		<td >	</td>		
		<td align="right" valign="Top">Ngày Xin</td>		
		<td align="right" valign="Top">Năm</td>		
		<td align="right" valign="Top">Tháng</td>		
		<td align="right" valign="Top">Ngày</td>		
	</tr>	
</table>	   
<table width="683" border="0" cellpadding="1" cellspacing="1" class="txt8" bgcolor="#000000" >
  <tr bgcolor="#ffffff" height=25>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr bgcolor="#ffffff">
    <td height="34">&nbsp;</td>
    <td colspan="4">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr bgcolor="#ffffff">
    <td width="80" height="34">&nbsp;</td>
    <td colspan="4">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr bgcolor="#ffffff">
    <td width="80">&nbsp;</td>
    <td colspan="7">&nbsp;</td>
  </tr>
  <tr bgcolor="#ffffff">
    <td width="80">日期<br/>Ngày</td>
    <td width="80">時間<br />Thời gian</td>
    <td width="80">說明<br />Lý d</td>
    <td width="80">交通費<br />Phí giao thông</td>
    <td width="80">繕雜費 <br />Tiền cơm</td>
    <td width="80">住宿費<br />Tiền nhà trọ</td>
    <td width="80">特別費<br />Phí khác</td>
    <td width="80">合計<br />Tổng cộng</td>
  </tr>
	<%for y = 1 to 5 %>
  <tr height=22 bgcolor="#ffffff">
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
  </tr>
	<%next%>
  </table> 
</body>
</html>
<!-- #include file="../Include/func.inc" -->

<script language=vbs>
function f()
	'<%=self%>.yymm.focus()
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
