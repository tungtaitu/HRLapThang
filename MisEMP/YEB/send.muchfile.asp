<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
self="sendMfile"  
'Set conn = GetSQLServerConnection()	  
 
if session("netuser")="" then 
	response.write "請先登入!!<BR>" 
	Response.Write "<b>UserID is empty or No limits, please Login again !!<br>Vao mang trong rong , hoac doi lau , hoac khong duoc su dong, <br>hay nhan nut nhap mang tu dau !!!</b>"
	response.end 
end if 
 
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">  
<title>YFYEMP</title> 
</head>
<body leftmargin="5"  topmargin="0"  marginwidth="0" marginheight="0" onload=f() onkeydown="enterto()">
<form name=<%=self%> method="post"  ENCTYPE="multipart/form-data" ACTION="sendmuchfile.send.asp"  >
<input type=hidden name=netuser value="<%=session("Netuser")%>">
<input type=hidden name=nowID value="<%=nowID%>">
<table   border=0 class=txt>
	<tr>
		<td align=left height=30   valign=bottom class=txt12  ><font color="#000099"><b>【外包員工照片】</b></font></td>
		<Td align=left class=txt12 valign=bottom width=400><b>批次檔案上傳<b></td> 
	</tr>
	<tr>
		<td align=center colspan=2   ><img src="../../vyfynet/picture/banner02.gif" width="550" height="15"></td>
	</tr>
</table>
<Table border=0  cellspacing="1" cellpadding="1"  class=txt9  bgcolor=#cccccc width=550> 
	<tr bgcolor=#ffffff>
	<td>STT</td> 
	<td>檔案</td> 
	</tr> 
	<%for x = 1 to 15%>
	<tr bgcolor=#ffffff>
		<Td align=center><%=x%></td> 
		<Td><INPUT TYPE="FILE" NAME="FILE<%=x%>" SIZE="60" class=inputbox></td>
	</tr>
	<%next%>
	<Tr bgcolor=#ffffff height=50>
		<td colspan=5 align=center>
			<input type=submit name=btn value="確認傳送" class=button >
			<input type=button name=btn value="關閉視窗" class=button onclick='parent.close()'>
		</td>
	</tr>
</table>    
</form>
</BODY>
</HTML>
<script language=vbscript> 

function f()
	'if <%=self%>.empid.value="" then 
	'	<%=self%>.empid.focus()
	'else
	'	<%=self%>.jb.focus()	 
	'end if 	
end function  

function goc()
	parent.close()
end function   


function empidchg(index)
	empid_str = <%=self%>.empid(index).value
	if empid_str<>"" then 
		open "sendMfile.back.asp?index="&index&"&empid="& empid_str , "Back" 
		'parent.best.cols="70%,30%"
	end if	
end function 

function sendtocq(index)
	writests = <%=self%>.writests(index).value
	khid = <%=self%>.khid(index).value
	if <%=self%>.writests(index).value="KH" then 
		if <%=self%>.tbi(index).value<>"Y" then 
			alert "(I)績效評核未填寫,無法送呈!!"
			exit function 
		end if
		if <%=self%>.tbii(index).value<>"Y" then 
			alert "(II)知識才能評核未填寫,無法送呈!!"
			exit function 
		end if
		if <%=self%>.tbiii(index).value<>"Y" then 
			alert "面談紀錄未填寫,無法送呈!!"
			exit function 
		end if		
		'if <%=self%>.nextKH(index).value<>"Y" then 
		'	alert "下半年度目標未填寫,無法送呈!!"
		'	exit function 
		'end if					
		
		open "sendcq.index.asp?writests=" & writests &"&khid=" & khid , "_blank" , "top=10, left=10, width=600, height=500, scrollbars=yes"
	else
		open "sendcq.index.asp?writests=" & writests &"&khid=" & khid , "_blank" , "top=10, left=10, width=600, height=500, scrollbars=yes"	
	end if		
end function 

 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function

function senfMulfile() 
	open "send.muchfile.asp" , "_blank" , "top=10, left=10, width=600, height=500, scrollbars=yes"
end function    
 
</script>


