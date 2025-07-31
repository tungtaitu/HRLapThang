<%option Explicit
Response.Buffer =True
%>
<!--#include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../Include/global_asp_fun.asp" -->
<%
const self		="YSBAE0401.UPDATEDB.ASP"
const action	="YSBAE0401.FORE.ASP"

const formname	="FRM"
const method	="POST"

dim conn,rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open GetSQLServerConnection()


'工作狀態，0為新增、1為修改及刪除
dim ActStatus,strSql
ActStatus = request.Form ("ActStatus")

dim SearchKey,GROUP_ID,PAGE
SearchKey	=request.Form ("SearchKey")
GROUP_ID	=request.Form ("GROUP_ID")
PAGE		=request.Form ("PAGE")

select case ActStatus
	case 0 'addnew
	case 1 'UpDate & Delete
		dim ChkDelete,Cnt,iRows
		dim Program_id,GROUP_R,GROUP_W
		Cnt = Request.Form ("Cnt")	'記載修改資料數量
		
		'呼叫SP_YSBAE0401_02來update
		for iRows = 0 to Cnt
			Program_id	= Request.Form ("Program_id"	& iRows)
			GROUP_R	= Request.Form ("GROUP_R"	& iRows)	
			GROUP_W	= Request.Form ("GROUP_W"	& iRows)	
			StrSql="SP_YSBAE0401_02 '"& trim(Program_id) &"','"& trim(GROUP_ID) &"','"& trim(GROUP_R) &"','"& trim(GROUP_W) &"'"
			'Response.Write strsql&"<br>"
			conn.Execute(StrSql)
		next
		conn.Close ()
		set conn=nothing
end select   
response.redirect "YSBAE0401.asp" 
%>
<html>
<head>
<script language =vbs>
function go()
	with window.<%=formname%> 
		.submit
	end with	
end function
</script>
</head>
<body background="..\..\Picture\bg9.gif" scroll=auto  onload ="go()">
<form NAME ="<%=formname%>" ACTION="<%=action%>" method="<%=method%>" id="<%=formname%>" >
	<INPUT type="hidden" value="<%=SearchKey%>" id=SearchKey name=SearchKey>
	<INPUT type="hidden" value="<%=GROUP_ID%>" id="GROUP_ID" name=GROUP_ID>
	<INPUT type="hidden" value="<%=PAGE%>" id="PAGE" name=PAGE>
</form>
</body>
</html>