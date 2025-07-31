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
  <tr height=25> 
    <td colspan=2 class=txt9C >E.員工請假作業</td>    
  </tr>
  <tr height=25> 
    <td width=8></td>    
    <td width=102>
    	<a href="employee/holiday/empholiday.asp" target="main">
    	1.請假作業新增</a>
    </td>    
  </tr>
  <tr height=25 > 
    <td width=8></td>    
    <td >
    	<a href="employee/holiday/empholidayB.asp" target="main">
    	2.請假作業維護</a>
    </td>  
  </tr>  
  <tr height=25 > 
    <td width=8></td>    
    <td >
    	<a href="employee/holiday/empholiday.fore.asp" target="main">
    	3.請假作業管理</a>
    </td>  
  </tr> 
  <tr> 
    <td colspan=2 align=left height=20></td>
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