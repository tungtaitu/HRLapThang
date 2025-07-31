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
    <td colspan=2 height=22 class=txt9C >F.報表資料查詢</td>    
  </tr>  
  <tr height=25 > 
    <td width=8></td>    
    <td height=25>
    	<img border="0" src="picture/icon01.gif" align="absmiddle" >
    	 <u>差勤資料報表</u> 
    </td>  
  </tr> 
  <tr height=25 >
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_basicCabiao.Fore.asp" target="main">
    	1.員工基本資料卡</a>
    </td>    
  </tr>
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_basicbiao.Fore.asp" target="main">
    	2.員工資料明細表</a>
    </td>    
  </tr>  
  <tr> 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_worktime.Fore.asp" target="main">
    	3.每日出勤明細表(依工號)</a>
    </td>    
  </tr>
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_worktime_bydate.fore.asp" target="main">
    	4. 每日出勤明細表(依日期)</a>
    </td>    
  </tr>  
   <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_worktimeN.fore.asp" target="main">
    	5. 員工出勤統計表</a>
    </td>    
  </tr>  
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_breakworktime.fore.asp" target="main">
    	6. 異常出勤紀錄明細表</a>
    </td>    
  </tr>   
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/emp_holiday.fore.asp" target="main">
    	7. 員工請假明細表</a>
    </td>    
  </tr>  
   <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/RenShiYD.fore.asp" target="main">
    	8. 人事異動申請表</a>
    </td>    
  </tr>   
  </tr>  
   <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/HOPDON.Fore.asp" target="main">
    	9. 海外工作人員合同</a>
    </td>    
  </tr> 
  <tr height=25 > 
    <td width=8></td>    
    <td width=102>
    	<a href="report/empworkTOT.Fore.asp" target="main">
    	A. 員工考核表</a>
    </td>    
  </tr>	
  <tr> 
    <td colspan=2 align=left height=20> </td>
  </tr> 
    
  
  <tr height=20 > 
    <td width=8></td>    
    <td >
    	<img border="0" src="picture/icon02.gif" align="absmiddle" >
    	<a href="employee/main.asp" target="main"><u>回主選單</u></a>
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