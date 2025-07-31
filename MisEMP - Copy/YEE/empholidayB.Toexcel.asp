<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "EMPHOLIDAYB"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
DAT1 = REQUEST("DAT1")
DAT2 = REQUEST("DAT2")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
country = trim(request("country"))  
QUERYX = trim(request("empid1"))  

unitno = trim(request("unitno"))
zuno = trim(request("zuno"))
job = trim(request("job")) 
jb = trim(request("jb")) 
ym1= trim(request("ym1")) 
ym2= trim(request("ym2")) 

if dat1="" and dat2="" and whsno="" and groupid="" and country="" and QUERYX="" then 
	sql="select * from empfile where empid='XX' "
else
	SQL="SELECT  A.JIATYPE,  CONVERT(CHAR(10), A.DATEUP, 111) DATEUP , A.TIMEUP, convert(char(10) , A.DATEDOWN , 111) datedown, "
	SQL=SQL&"A.TIMEDOWN , A.HHOUR, A.MEMO AS JIAMEMO  , a.autoid as jiaid,  B.*  , isnull(c.sys_value,'') as jia_str  FROM   "
	SQL=SQL&"( SELECT * FROM EMPHOLIDAY   ) A  "
	SQL=SQL&"LEFT JOIN ( SELECT * FROM view_empfile ) B ON B.EMPID = A.EMPID  	 "
	SQL=SQL&"LEFT JOIN ( SELECT * FROM basicCode where func='JB'  ) c  on c.sys_type = a.JIATYPE  "
	SQL=SQL&"WHERE 1=1  " 	
	SQL=SQL&"and country like  '"& country &"%'  "
	SQL=SQL&"AND whsno like '"& whsno &"%' and unitno like '"& unitno &"%'  and groupid like '"& groupid &"%'  " 
	SQL=SQL&"and zuno like '"& zuno &"%' and a.jiatype like '"& jb &"%' and b.empid like '%"& QUERYX &"%'  "
	
	IF DAT1<>"" and DAT2<>"" then  
	 	sql=sql& "and CONVERT(CHAR(10), A.DATEUP, 111) BETWEEN '"& DAT1 &"' AND '"& DAT2 &"' " 
	END IF 
	IF ym1<>"" and ym2<>"" then  
	 	sql=sql& "and CONVERT(CHAR(6), A.DATEUP, 112) BETWEEN '"& ym1 &"' AND '"& ym2 &"' " 
	END IF 
	SQL=SQL&"order by b.empid, A.DATEUP , a.jiaType "  
end if 	
'response.write sql 
'RESPONSE.END  
rs.Open SQL, conn, 3, 3 

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<%
  filenamestr = "empholiday.xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
%>
<table width=100% class="txt8"  cellspacing="1" cellpadding="1"  >
	<tr BGCOLOR="LightGrey" height=22>
		<TD width=50 nowrap align=center >工號<br>So The</TD> 		
 		<TD width=190 nowrap align=center >姓名<br>Ho ten</TD>
 		<TD align=center  >假別<br>loai phep</TD>
 		<TD width=80 align=center nowrap >日期(起)<br>Ngay(tu)</TD>
		<TD align=center  >時間(起)<br>Thoi gian(tu)</TD>
		<TD width=80 align=center nowrap >日期(迄)<br>Ngay(Den)</TD>
		<td align=center  >時間(迄)<br>Thoi gian(den)</td>
		<td align=center  >時數<br>So gio</td>
		<td align=center >事由<br>Ly do</td>		
	</tr>
	<%
	while not rs.eof 	
		 
	%>	
	<TR > 
		<TD align=center>
			<%=trim(rs("empid"))%>
		</TD> 		
 		<TD>
 			<%=trim(rs("empnam_cn"))%>&nbsp;
 			<font class=txt8><%=trim(rs("empnam_vn"))%></font>
 		</TD>
 		<TD>
 			<%=RS("JIATYPE")%>&nbsp;<%=RS("jia_str")%>
 		</TD>
 		<TD align=center>
 			<%=RS("DATEUP") &" "&mid("日一二三四五六",weekday(cdate(rs("DATEUP"))) , 1 )%>
 		</TD>
 		<TD align=center>
 			<%=RS("TIMEUP")%>
 		</TD>
 		<TD align=center>
 			<%=RS("DATEDOWN") &" "&mid("日一二三四五六",weekday(cdate(rs("DATEDOWN"))) , 1 )%>
 		</TD> 
 		<TD align=center>
 			<%=RS("TIMEDOWN")%>
 		</TD>
 		<TD align=center>
 			<%=RS("hhour")%>
 		</TD> 
 		<TD align=center>
 			<%=RS("JIAMEMO")%>
 		</TD> 		
	</TR>
<%
rs.movenext
wend 
set rs=nothing 
%>	
</table>
</form>
</body>
</html>