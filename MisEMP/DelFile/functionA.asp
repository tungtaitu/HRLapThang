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
  <tr height=25 >
    <td colspan=2 class=txt9C>A.�򥻸�ƫ���</td>
  </tr>
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="admin/admin01.asp" target="main">1.�N�X�ɺ��@</a></td>
  </tr>
	<%if session("netuser")="PELIN" then %>
	<tr height=25 >
	    <td ></td>
	    <td ><a href="employee/empbasic/YSBHE0501.asp" target="main">2.�s�W�ϥΪ�</td>
	</tr>  
	<tr height=25>
		<td width=8></td>
	    <td width=102><a href="syspro/ysbae0101.asp" target="main">3.�{���\����@</a></td>
	</tr>
  <%end if %> 
  <tr height=25 >
    <td ></td>
    <td ><a href="employee/empbasic/YSBHE0501.asp" target="main">A.�ק�K�X</td>
  </tr> 
  <tr height=25>
    <td width=8></td>
    <td width=102><a href="employee/empbasic/empbasic.asp" target="main">B.�򥻫���</a></td>
  </tr>
  <tr height=25 >
    <td ></td>
    <td ><a href="employee/empbasic/empbasicB.Fore.asp" target="main">C.�𰲤�]�w</td>
  </tr>
  <tr height=25 >
    <td ></td>
    <td ><a href="employee/empbasic/YFYEXRT.asp" target="main">D.�ײv</td>
  </tr>
 
  <tr>
    <td colspan=2 align=left height=20> </td>
  </tr>


 <tr height=25 >
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