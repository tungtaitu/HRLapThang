<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%
Set CONN = GetSQLServerConnection()
Set rds = Server.CreateObject("ADODB.Recordset")      
Set rs2 = Server.CreateObject("ADODB.Recordset")      

sqln="select  a.* , b.empnam_cn  from "&_
	 "( select   empid, max(edat) edat  from  empvisadata  group by  empid  ) a   "&_
	 "join  ( select empid , empnam_cn, empnam_vn , outdat from view_empfile ) b  on b.empid = a.empid  "&_
	 "where  isnull(b.outdat,'')=''    "&_
	 "and  convert(char(10),a.edat,111) < convert(char(10), dateadd( d, 20, getdate()) , 111)  "&_
	 "order by   convert(char(10),(edat),111) " 
rds.open sqln , conn, 3, 3   
'response.write sqln 
sqld="select  a.* , b.empnam_cn  from "&_
	 "( select   empid, max(edat) edat  from  empHTdata  group by  empid  ) a   "&_
	 "join  ( select empid , empnam_cn, empnam_vn , outdat from view_empfile ) b  on b.empid = a.empid  "&_
	 "where  isnull(b.outdat,'')=''    "&_
	 "and  convert(char(10),a.edat,111) < convert(char(10), dateadd( d, 20, getdate()) , 111)  "&_
	 "order by   convert(char(10),(edat),111) "  
rs2.open sqld , conn, 3, 3   
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<link rel="stylesheet" href="../Include/bar_v3.css" type="text/css">
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0" >
<table width="460" border="0" cellspacing="0" cellpadding="0" >
	<tr><TD>人事薪資系統</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<BR>
<table width=500><tr><td>
	<div id="navcontainer">
<ul id="navlist"> 
	<table width=400 align=center cellspacing="3" cellpadding="3" >	
		<tr height=50>
			<td height=25 align=left valign="top"> 
				<li  ><a href="vbscript:dataclick(1)" id="A1"><b>A.基本資料建檔</b><br><font class=txt>TẠO TƯ LIỆU CƠ BẢN</font></a></li>
			</td>
			<td height=25 align=left valign="top"> 
				<li  ><a href="vbscript:dataclick(2)"><b>B.員工基本資料</b><br><font class=txt>TƯ LIỆU CƠ BẢN CỦA NHÂN VIÊN</font></a></li>
			</td> 
		</tr> 
		<tr height=50>
			<td height=25 align=left valign="top"> 
				<li><a href="vbscript:dataclick(7)"><b>C.薪資管理作業</b><br><font class=txt>QUẢN LÝ TIỀN LƯƠNG</font></a></li>
			</td>
			<td height=25 align=left valign="top"> 
				<li><a href="vbscript:dataclick(3)"><b>D.員工差勤作業</b><br><font class=txt>CHẤM CÔNG NHÂN VIÊN</font></a></li>
			</td>
		</tr>
		<tr height=50 >
			<td height=25 align=left valign="top"> 
				<li><a href="vbscript:dataclick(4)"><b>E.員工請假作業</b><br><font class=txt>TƯ LIỆU NHÂN VIÊN NGHỈ PHÉP</font></a></li>
			</td>
			<td height=25 align=left valign="top"> 
				<li><a href="vbscript:dataclick(6)"><b>F.報表資料管理</b><br><font class=txt>BẢO VỆ BẢNG BIỂU TƯ LIỆU</font></a></li>
			</td>
		</tr>	
		<tr height=50 >
			<td height=25 align=left valign="top"> 
				<li><a href="vbscript:dataclick(5)"><b>G.接收卡鐘資料</b><br><font class=txt>TIẾP THU TƯ LIỆU MÁY GẠT THẺ</font></a></li>
			</td>
			<td height=25 align=left valign="top"> 
				<li><a href="vbscript:dataclick(8)"><b>H.員工扣款作業</b><br><font class=txt>KHẤU TRỪ LƯƠNG NHÂN VIÊN</font></a></li>
			</td>
		</tr>
		<tr height=50 >
			<td height=25 align=left valign="top"> 
				<li><a href="vbscript:dataclick(9)"><b>I.員工績效考核</b><br><font class=txt>KHẢO HẠCH NĂNG SUẤT NHÂN VIÊN</font></a></li>
			</td>
			<td height=25 align=left>			
			</td>
		</tr>		
	</table>  
</ul>
</div>
	<BR><BR>
	
	 
	



</td></tr></table>

</body>
</html>


<script language=vbs>
	function dataclick(a)
		if a = 1 then			
			open "../function.asp?program_id=A" , "contents"
		elseif a = 2 then			
			'open "../functionB.asp" , "contents"
			open "../function.asp?program_id=B" , "contents"
		elseif a = 3 then
			'open "empworkHour/empwork.asp" , "_self"
			'open "../functionD.asp" , "contents"
			open "../function.asp?program_id=D" , "contents"
		elseif a = 4 then
			'open "holiday/empholiday.asp" , "_self"
			'open "../functionE.asp" , "contents"
			open "../function.asp?program_id=E" , "contents"
		elseif a = 5 then
			'open "AcceptCaTime/main.asp" , "_self"
			'open "../functionG.asp" , "contents"
			open "../function.asp?program_id=G" , "contents"
		elseif a = 6 then
			'open "../report/main.asp" , "_self"
			'open "../functionF.asp" , "contents"
			open "../function.asp?program_id=F" , "contents"
		elseif a = 7 then
			'open "../functionC.asp" , "contents"
			open "../function.asp?program_id=C" , "contents"
		elseif a = 8 then
			'open "../functionH.asp" , "contents"
			open "../function.asp?program_id=H" , "contents"
		elseif a = 9 then
			'open "../functionH.asp" , "contents"
			open "../function.asp?program_id=I" , "contents"
		end if
	end function
</script>