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
	<tr><TD>�����Ƭd��</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<BR>
<table width=500  ><tr><td  >
	<table width=450 align=center border=0 cellspacing="0" cellpadding="0"  >
		<tr height=50>
			<td align=center height="35"><img border="0" src="../Picture/icon02.gif" align="absmiddle"> 
              �H�Ʈt�Ը�Ƴ���</td>
			<td align=center height="35" ><img border="0" src="../Picture/icon02.gif" align="absmiddle"> 
			 ���u�~���Ƴ���</td>
		</tr>
		<tr height=50>
			<td align=center valign=top >
				<table width=140 align=center border=0 cellspacing="0" cellpadding="0" class=font9 height="30"> 
					<tr><td  height="22"><a href="emp_basicCabiao.getrpt.asp"><u>1. ���u�򥻸�ƥd</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_basicbiao.Fore.asp"><u>2. ���u��Ʃ��Ӫ�</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_worktime.fore.asp"><u>3. ���u�t�Ԭ������Ӫ�(�̤u��)</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_worktime_bydate.fore.asp"><u>4. ���u�C��X�ԩ��Ӫ�(�̤��)</u></a></td></tr>  
					<tr><td  height="22"><a href="emp_Breakworktime.fore.asp"><u>5. ���u���`�X�Ԭ������Ӫ�</u></a></td></tr>					  
				</table>
			</td>
			<td align=center valign=top  >
				<table width=130 align=center border=0 cellspacing="0" cellpadding="0" class=font9> 
					<tr><td  height="22"><a href=""><u>1. ���u�򥻸�ƥd</u></a></td></tr>  
					<tr><td  height="22"><a href=""><u>2. ���u�򥻸�Ʃ��Ӫ�</u></a></td></tr>
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