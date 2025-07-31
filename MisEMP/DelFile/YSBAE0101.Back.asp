<%@Language=VBScript codepage=65001%>
<!--#include file="../GetSQLServerConnection.fun"--> 
<%

Response.Expires = 0	
%>
<HTML>
<HEAD>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta http-equiv="refresh">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
<BODY> 
<form name=form1 >
LoginType <input type=text name="LoginType" value ="<%=session("LoginType") %>">
</form>
<%
func = request("func") 
index = request("index")
CurrentPage = request("CurrentPage")  
pid = trim(request("pid"))
program_id = request("program_id")
program_name = request("program_name")



layer_up = request("layer_up")
layer = request("layer")
PROGRAM_NAME_VN = request.QueryString("PROGRAM_NAME_VN")
VIRTUAL_PATH = request("VIRTUAL_PATH")  

proname = Server.HTMLEncode(Request("program_name"))  

response.write "1=" & request.QueryString("program_name") &"<BR>"
response.write "2=" & proname &"<BR>"  
response.write  "3=" & program_name &"<BR>"
response.write "4=" &  server.URLEncode(request.QueryString("program_name")) &"<BR>"


response.write program_id &"<BR>"
response.write program_name &"<BR>"
response.write layer_up &"<BR>"
response.write layer &"<BR>"
response.write PROGRAM_NAME_VN &"<BR>"
response.write VIRTUAL_PATH &"<BR>"
 
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open GetSQLServerConnection()

tmpRec = Session("syspro01")

Select Case func
	   case "pidchg" 
	   	sql = "select * from SYSPROGRAM where program_id='"& pid  &"' "
	   	'response.write sql 
	   	Set rs = Server.CreateObject("ADODB.Recordset") 
	   	rs.open sql, conn, 3, 3 
	   	if not rs.eof then 
%>			<script language=vbscript>
				parent.Fore.FRM.pname.value="<%=rs("program_name")%>"
				parent.Fore.FRM.pnameVN.value="<%=rs("proname_VN")%>"
				parent.Fore.FRM.UPID.value="<%=rs("layer_up")%>"
				parent.Fore.FRM.LEVEL.value="<%=rs("layer")%>"
				parent.Fore.FRM.vpath.value="<%=rs("VIRTUAL_PATH")%>"
				ltstr=form1.LoginType.value 
				if ltstr<>"AAA" then 
					parent.Fore.FRM.pnameVN.focus()
				end if 
			</script>
<%	   	else%>
			<script language=vbscript>
				parent.Fore.FRM.pname.value=""
				parent.Fore.FRM.pnameVN.value=""
				parent.Fore.FRM.UPID.value=""
				parent.Fore.FRM.LEVEL.value=""
				parent.Fore.FRM.vpath.value=""
			</script>
<%			set rs=nothing 
		end if 
	   Case "datachg"
			tmpRec(CurrentPage,index + 1,0) = "upd"
			tmpRec(CurrentPage,index + 1,1) = program_id
			tmpRec(CurrentPage,index + 1,2) = program_name
			tmpRec(CurrentPage,index + 1,3) = layer_up
			tmpRec(CurrentPage,index + 1,4) = layer
			tmpRec(CurrentPage,index + 1,5) = VIRTUAL_PATH
			tmpRec(CurrentPage,index + 1,6) = PROGRAM_NAME_VN			
				
	   Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del" 
	   Case "no"
			tmpRec(CurrentPage,index + 1,0) = "upd" 			
	
End Select
response.write "0_=" & tmpRec(CurrentPage,index + 1,0) &"<BR>"
response.write "1_=" & tmpRec(CurrentPage,index + 1,1) &"<BR>" 
response.write "2_=" & tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write "3_=" & tmpRec(CurrentPage,index + 1,3) &"<BR>" 
response.write "4_=" & tmpRec(CurrentPage,index + 1,4) &"<BR>" 
response.write "5_=" & tmpRec(CurrentPage,index + 1,5) &"<BR>" 
response.write "6_=" & tmpRec(CurrentPage,index + 1,6) &"<BR>" 
response.write "7_=" & tmpRec(CurrentPage,index + 1,7) &"<BR>" 

Session("syspro01") = tmpRec


%> 
<form name=form1 >
	<input name=t1 value="<%=program_name%>">
	<input name=t1 value="<%=server.urlencode(PROGRAM_NAME_VN)%>">
	<input name=t1 value="<%=server.HTMLencode(PROGRAM_NAME_VN)%>">
</form> 
</BODY>
</HTML>


