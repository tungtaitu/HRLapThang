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
    <td colspan=2 class=txt9C >C.薪資管理作業</td>
  </tr>
  <%if session("netuser")="PELIN" OR session("netuser")="L0051" or session("netuser")="L5002"   or session("netuser")="L0627"  then %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/empsalary01.asp" target="main">
    1.基本薪資管理</a></td>
  </tr>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/empsalary.asp" target="main">
    2.員工薪資計算</a></td>
  </tr>
   <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/empsalaryHW.asp" target="main">
    3.泰國員工薪資</a></td>
  </tr>
  <%end if %>
  </tr>
  <%if session("netuser")="PELIN" or session("netuser")="L5002" THEN %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/empsalaryCN.asp" target="main">
    4.中國員工薪資</a></td>
  </tr>
  <%END IF %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/YFYEMPJX.asp" target="main">
    5.當月績效</a></td>
  </tr>
  <%if session("netuser")="PELIN" OR session("netuser")="L0051" or session("netuser")="L5002"  or session("netuser")="L0627"  then %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/YFYEMPJXA.asp" target="main">
    6.績效獎金</a></td>
  </tr>
  <%end if %>
  <tr>
    <td colspan=2 align=left height=20> </td>
  </tr>
   <tr height=25>
    <td colspan=2 class=txt9C >C.報表列印作業</td>
  </tr>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp01.asp" target="main">
    1.員工名條列印</a></td>
  </tr>
  <%if session("netuser")="PELIN" OR session("netuser")="L0051" or session("netuser")="L5002"  or session("netuser")="L0627"   then %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp02.asp" target="main">
    2.員工薪資單列印</a></td>
  </tr>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp03.asp" target="main">
    3.薪資明細表</a></td>
  </tr>
   <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp031.asp" target="main">
    3.1 年終獎金明細</a></td>
  </tr>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp04.asp" target="main">
    4.薪資彙總表</a></td>
  </tr>
  <%end if %>
  <%if session("netuser")="PELIN" OR session("netuser")="L0051"  or session("netuser")="L5002"   or session("netuser")="L0627"  then %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp05.asp" target="main">
    5.員工加班明細表</a></td>
  </tr>
  <%end if %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp06.asp" target="main">
    6.員工薪資簽收表</a></td>
  </tr>
  <%if session("netuser")="PELIN" OR session("netuser")="L0051"  or session("netuser")="L5002"   or session("netuser")="L0627"  then %>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp07.asp" target="main">
    7.基本明細表</a></td>
  </tr>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empsalary/salarycp08.asp" target="main">
    8.績效獎金明細表</a></td>
  </tr>
  <%end if %>
  <tr height=25>
    <td width=8></td>
    <td width=102> </td>
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