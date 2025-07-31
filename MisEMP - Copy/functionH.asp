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
    <td colspan=2 height=22 class=txt9C >H.員工扣款作業</td>
  </tr>
  <tr height=25>
  	<td width=8></td>
    <td width=102>
    	<a href="yfysuco/VYFYSUCO1.asp" target="main">
    	1.轉入事故扣款</a>
    </td>
  </tr>
  <tr height=25 >    
    <td width=8></td>
    <td width=102>
    	<a href="yfysuco/VYFYSUCO.asp" target="main">
    	2.扣款新增作業</a>
    </td>
  </tr>
  <tr height=25 >
    <td width=8></td>
    <td width=102>
    	<a href="yfysuco/VYFYSUCOS.asp" target="main">
    	3.事故(扣款)查詢</a>
    </td>
  </tr>


  <tr>
    <td colspan=2 align=left height=20> </td>
  </tr>
  </tr>
   <tr height=25>
    <td colspan=2 class=txt9C >C.報表列印作業</td>
  </tr> 
  <tr height=25 >
    <td width=8></td>
    <td width=102>
    	<a href="yfysuco/VYFYSUCOS01.asp" target="main">
    	1.事故(扣款)明細</a>
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