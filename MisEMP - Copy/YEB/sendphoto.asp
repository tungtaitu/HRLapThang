<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
self="sendMfile"  
Set conn = GetSQLServerConnection()	  
 
if session("netuser")="" then 
	response.write "請先登入!!<BR>" 
	Response.Write "<b>UserID is empty or No limits, please Login again !!<br>Vao mang trong rong , hoac doi lau , hoac khong duoc su dong, <br>hay nhan nut nhap mang tu dau !!!</b>"
	response.end 
end if 

empid = request("empid")

sqln="select * from empfile where empid='"& request("empid") &"' "
set rst=conn.execute(sqln) 
if not rst.eof then 
	empname=rst("empnam_cn")
end if 
set rst=nothing
flag=request("flag")
conn.close 
set conn=nothing
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
 
<title>VYFYNET</title>
 
</head>
<body leftmargin="5"  topmargin="10"  marginwidth="0" marginheight="0" onload=f() onkeydown="enterto()">
<form name=<%=self%> method="post"  ENCTYPE="multipart/form-data"    >
<input type=hidden name=netuser value="<%=session("Netuser")%>">
<input type=hidden name=nowID value="<%=nowID%>">
<input type=hidden name=EMPID value="<%=EMPID%>">
<input type=hidden name=flag value="<%=flag%>">

<table width="250" border="0"   cellspacing="2" cellpadding="1" >
	<tr> 
		<td>工號姓名 : <%=empid%><%=empname%></td>
	</tr>	
	<tr> 
		<td align=left > -- 上傳照片(photo) --</td>
	</tr>				
	<tr>
		<Td>
			<INPUT TYPE="FILE" NAME="FILE1" SIZE="40" class=inputbox8 >
		</td>
	</tr>			
	<tr> 
		<td align=left > -- 上傳護照(passport) --</td>
	</tr>				
	<tr>
		<Td>
			<INPUT TYPE="FILE" NAME="FILE2" SIZE="40" class=inputbox8 >
		</td>
	</tr>			
	<tr> 
		<td align=left > -- 上傳簽証工作証(visa) --</td>
	</tr>				
	<tr>
		<Td>
			<INPUT TYPE="FILE" NAME="FILE3" SIZE="40" class=inputbox8 >
		</td>
	</tr>			
	<Tr bgcolor=#ffffff height=50>
		<td colspan=5 align=center>
			<input type=button name=btn value="(Y)Confirm" class=button onclick=go()>
			<input type=button name=btn value="(X)Close" class=button onclick='parent.close()'>
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

function go()
	
	if trim(<%=self%>.file1.value)<>"" and Ucase(right(<%=self%>.file1.value,3))<>"JPG" then 
		alert "格式錯誤(需為jpg檔)!! Wrong Type (*.jpg)"
		<%=self%>.file1.focus()
		exit function 
	end if 	
	if <%=self%>.flag.value="WB" then 
		<%=self%>.action="send2wbphoto.asp"
		<%=self%>.submit()
	else
		<%=self%>.action="send2.asp"
		<%=self%>.submit()
	end if 
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


