<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
'on error resume next
session.codepage="65001"
SELF = "YEBE0301"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
job = trim(request("job"))
country = request("country")
QUERYX = trim(request("empid1"))
outemp = request("outemp")
EMPID = REQUEST("empid")
shift = request("shift")
IOemp = request("IOemp")
inym = request("inym")
zuno = request("zuno")

NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)


sql="select * from fn_View_EMPFILE ('"& whsno &"','"& country&"','"& QUERYX &"' , '"& inym &"', '"& groupid &"','"&zuno &"', '"&shift&"','"&IOemp &"' ) where 1=1  "
if IOemp="Y" then
		sql = sql & " AND ( ISNULL(OUTDAT,'')='' OR convert( char(6),OUTDAT,111)>='"& NOWMONTH &"' )  "		
	elseif IOemp="N" then
		sql = sql & " AND ( ISNULL(OUTDAT,'')<>'' )  "
	end if
	if inym<>"" then
		sql = sql & " and convert(char(6),indat,112)=   '"& inym &"'  "
	end if 
sql = sql & "order by empid  "
 
rs.Open sql, conn, 3, 1
'response.write sql 
'response.end
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<style >
.txtvn8 { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 8pt }  
.txtvn9 { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 10pt }
.txtTotal { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 12pt }
.txtvn14 { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 14pt }  
</style> 
</head> 
<body  > 
<%
  ilenamestr = "ThongTinNhanVien.xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel" 
%>
<TABLE style="height:50px" class="txtvn14" BORDER=0 cellspacing="3" cellpadding="3" >	
	<tr>
		<td colspan=4 align="center">CÔNG TY TNHN HOÀ ĐƯỜNG<br>和唐責任有限公司</td>
		<td colspan=8 align="center"></td>
	</tr>	
</table>

<TABLE CLASS="txtvn8" BORDER=1 cellspacing="1" cellpadding="1" >
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>STT</td>				
		<td align=center>工號<br>Số thẻ</td>
		<td align=center>姓名<br>Họ Tên</td>
		<td align=center>部門<br>Bộ Phận</td>
		<td align=center>出生日期<br>Ngày sinh</td>
		<td align=center>到職<br>Vào xưởng</td>			
		<td align=center>簽合同<br>Ký hợp đồng</td>
		<td align=center>離職日<br>Thôi việc</td>
		<td align=center>身分證號<br>Số CMND</td>
		<td align=center>稅號<br>Mã số thuế</td>
		<td align=center>銀行帳號<br>Số tài khoản</td>
		<td align=center>保險號碼<br>Số Bảo hiểm</td>
	</tr>
	<%
	
	x = 0 
	  grp_cnt = 0	
	  grp_amt = 0 
	  grp_id = ""
	  tot_amt= 0 
	  while not rs.eof   
	  x=x+1
	%> 	

		<TR HEIGHT=22 BGCOLOR="#ffffff" class="txtvn9">					
			<td align=left><%=x%></td>			
			<td align=left><%=rs("empid")%></td>
			<td align=left nowrap><%=rs("empnam_cn")&rs("empnam_vn")%></td>	
			<td align=left><%=rs("gstr")%></td>
			<td align=left><%=rs("bdy_ymd")%></td>
			<td align=left><%=rs("nindat")%></td>			
			<td align=left><%=rs("bhdat")%></td>
			<td align=left><%=rs("outdate")%></td>
			<td align=left><%=rs("PERSONID")%></td>
			<td align=left><%=rs("taxCode")%></td>			
			<td align=left>&nbsp;<%=rs("BANKID")%></td>
			<td align=left><%=rs("soBH")%></td>
		</tr> 
	<%	 
	rs.movenext
	%> 
	<%wend%>  
	
</table> 
<%
rs.close
set rs=nothing 
conn.close 
set conn=nothing 
response.end%>

</body>
</html> 