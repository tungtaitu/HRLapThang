<%CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="yece12"  
txt_istr=""

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="1" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="1" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 
 
 
yymm=request("yymm") : if request("yymm") <>"" then calcmonth = request("yymm") else calcmonth=calcmonth
whsno=request("whsno")
'  sp_api_cb05 @yymm as varchar(6) , @CT as varchar(50), @whsno as varchar(5) ,@groupid as varchar(10) as 
If Request.ServerVariables("REQUEST_METHOD") = "POST" then 
	yymm=request("yymm")
	sql="exec sp_api_cb05 '"&yymm&"','VN','"&whsno&"', ''  "
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 1	
	if not rs.eof then 
		txt_istr ="" 
		
		set f1=server.createobject("scripting.filesystemobject")
		set temp=f1.createtextfile(server.mappath("files/transfile_"&yymm&".txt"),true,false)		
		while not rs.eof
			lens = 0 : txtSalary = "" 
			for i = 1 to 53 
				txt_istr = txt_istr & rs(i) 			
				txtSalary  = txtSalary & rs(i) 			 
				lens = lens+len(rs(i)) 				
			next 							
			'response.write  txtSalary &"<BR>"
			temp.WriteLine(txtSalary)
			txt_istr = txt_istr & " ---- "&lens&"<BR>"
			
			rs.movenext
		wend  
		'response.write txt_istr  
		temp.Close 
	end if	
	rs.close :set rs=nothing 
	'conn.close : set conn=nothing 
end if

 

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.YYMM.focus()	
	<%=self%>.YYMM.SELECT()
end function    
-->
</SCRIPT>   

</head>
	
<script  type="text/javascript">
function go(){
	var m = document.forms[0];
	m.action ="yecb05.fore.asp?pgid=<%=request("pgid")%>" ;
	m.submit();
}
</script>	 
<body   onkeydown="enterto()" onload="document.getElementById('yymm').select()"  >
<form name="<%=self%>" method="post" action="yecb05.ASP?pgid=<%=request("pgid")%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table class="txt" cellpadding=3 cellspacing=3> 
					<tr >
						<TD nowrap align=right>計薪年月<br>Tien luong</TD>
						<TD ><INPUT type="text" style="width:100px" NAME="yymm" id="yymm" VALUE="<%=calcmonth%>"></TD>							 
						<TD nowrap align=right height=30 >廠別<br>Xuong</TD>
						<TD > 
							<select name=whsno style="width:100px"  >
								<option value="">--ALL---</option>
								<%
								if  request("whsno")="" then whsno=session("mywhsno") else whsno=request("whsno")
								SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if whsno=RST("SYS_TYPE") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								%>					
							</SELECT>
							<%rst.close : SET RST=NOTHING %>
						</TD>					
						<td>
							<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
						</td>
					</tr>					
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table border=0 cellpadding=3 cellspacing=3 width="98%">
					<tr>
						<td>
							<%If Request.ServerVariables("REQUEST_METHOD") = "POST" then %>
							<table width="100%" class="txt8">
							<tr><td>結果 <a href="files/transfile_<%=yymm%>.txt" style="margin-left:10px;color:blue;font-size:12pt">[下載文字檔](滑鼠右鍵選擇另存目標)</a></td></tr>
							<tr><td><%=txt_istr%></td></tr>
							</table>
							<%end if%>	

							<%conn.close : set conn=nothing%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
			
</form>

</body>
</html>

 