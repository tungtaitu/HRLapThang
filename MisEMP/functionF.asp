<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "GetSQLServerConnection.fun" -->

<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=BIG5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	
<link rel="stylesheet" href="Include/style.css" type="text/css">
<link rel="stylesheet" href="Include/style2.css" type="text/css">
</head>
<body  topmargin="20" leftmargin="5"  marginwidth="0" marginheight="0">
<table width="110" border="0" cellpadding="0" cellspacing="0"  class=font9 align=center>  
  <tr> 
    <td colspan=2 height=22 class=txt9C >F.�����Ƭd��</td>    
  </tr>  
  <tr height=25 > 
    <td width=8></td>    
    <td height=25>
    	<img border="0" src="picture/icon01.gif" align="absmiddle" >
    	 <u>�t�Ը�Ƴ���</u> 
    </td>  
  </tr> 
  <tr height=25 >
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_basicCabiao.Fore.asp" target="main">
    	1.���u�򥻸�ƥd</a>
    </td>    
  </tr>
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_basicbiao.Fore.asp" target="main">
    	2.���u��Ʃ��Ӫ�</a>
    </td>    
  </tr>  
  <tr> 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_worktime.Fore.asp" target="main">
    	3.�C��X�ԩ��Ӫ�(�̤u��)</a>
    </td>    
  </tr>
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_worktime_bydate.fore.asp" target="main">
    	4. �C��X�ԩ��Ӫ�(�̤��)</a>
    </td>    
  </tr>  
   <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_worktimeN.fore.asp" target="main">
    	5. ���u�X�Բέp��</a>
    </td>    
  </tr>  
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_breakworktime.fore.asp" target="main">
    	6. ���`�X�Ԭ������Ӫ�</a>
    </td>    
  </tr>   
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_holiday.fore.asp" target="main">
    	7. ���u�а����Ӫ�</a>
    </td>    
  </tr>  
   <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/RenShiYD.fore.asp" target="main">
    	8. �H�Ʋ��ʥӽЪ�</a>
    </td>    
  </tr>   
  </tr>  
   <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/HOPDON.Fore.asp" target="main">
    	9. ���~�u�@�H���X�P</a>
    </td>    
  </tr> 
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/empworkTOT.Fore.asp" target="main">
    	A. ���u�Ү֪�</a>
    </td>    
  </tr>	
  <tr> 
    <td colspan=2 align=left height=20> </td>
  </tr> 
    
  
  <tr height=20 > 
    <td width=8></td>    
    <td >
    	<img border="0" src="picture/icon02.gif" align="absmiddle" >
    	<a href="employee/main.asp" target="main"><u>�^�D���</u></a>
    </td>  
  </tr>  
   
</table>

</body>
</html> 


<script language=vbs>
function BACKMAIN() 	
	open "employee/main.asp" , "main"
end function  
</script> 