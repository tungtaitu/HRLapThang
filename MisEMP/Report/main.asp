<%@LANGUAGE=VBSCRIPT CODEPAGE=950%>
<%
 
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0" >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>報表資料查詢</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<BR>
<table width=500  ><tr><td  >
	<table width=450 align=center border=0 cellspacing="0" cellpadding="0"  >
		<tr height=50>
			<td align=center height="35"><img border="0" src="../Picture/icon02.gif" align="absmiddle"> 
              人事差勤資料報表</td>
			<td align=center height="35" ><img border="0" src="../Picture/icon02.gif" align="absmiddle"> 
			 員工薪資資料報表</td>
		</tr>
		<tr height=50>
			<td align=center valign=top >
				<table width=140 align=center border=0 cellspacing="0" cellpadding="0" class=font9 height="30"> 
					<tr><td  height="22"><a href="emp_basicCabiao.getrpt.asp"><u>1. 員工基本資料卡</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_basicbiao.Fore.asp"><u>2. 員工資料明細表</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_worktime.fore.asp"><u>3. 員工差勤紀錄明細表(依工號)</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_worktime_bydate.fore.asp"><u>4. 員工每日出勤明細表(依日期)</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_Breakworktime.fore.asp"><u>5. 員工異常出勤紀錄明細表</u></a></td></tr>					  
				</table>
			</td>
			<td align=center valign=top  >
				<table width=130 align=center border=0 cellspacing="0" cellpadding="0" class=font9> 
					<tr><td  height="22"><a href=""><u>1. 員工基本資料卡</u></a></td></tr>  
					<tr><td  height="22"><a href=""><u>2. 員工基本資料明細表</u></a></td></tr>
				</table>			
			</td>
		</tr>
		 
	</table>		
	

</td></tr></table> 

</body>
</html>


<script language=vbs>
	function dataclick(a)
		if a = 1 then 		
			open "empbasic/empbasic.asp" , "_self"
		elseif a = 2 then 		
			open "empfile/empfile.asp" , "_self"
		elseif a = 3 then 		
			open "empworkHour/empwork.asp" , "_self"	
		elseif a = 4 then 		
			open "holiday/empholiday.asp" , "_self"	
		elseif a = 5 then 		
			open "AcceptCaTime/main.asp" , "_self"				
		elseif a = 6 then 		
			open "../report/main.asp" , "_self"		
		end if 		
	end function 
</script>