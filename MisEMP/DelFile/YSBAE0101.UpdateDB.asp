<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache"> 
</head>
<%
Response.Buffer =True
%>
<!-------- #include file = "../GetSQLServerConnection.fun" --------->
<!-- #include file="../Include/global_asp_fun.asp" -->
<%
const self		="YSBAE0101.UPDATEDB.ASP"
const action	="YSBAE0101.FORE.ASP"

const formname	="FRM"
const method	="POST"

dim conn,rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open GetSQLServerConnection()


'工作狀態，0為新增、1為修改及刪除
dim ActStatus,strSql
ActStatus = request.Form ("ActStatus")

'dim COMNO,COMNAME,ChkDelete,TxtCOMNO,TxtCOMNAME

select case ActStatus
	case 0 'addnew
		dim PROGRAM_ID,PROGRAM_NAME,LAYER_UP,LAYER,VIRTUAL_PATH
		PROGRAM_ID		= request.Form ("PROGRAM_ID")	'程式代碼
		PROGRAM_NAME	= request.Form ("PROGRAM_NAME")	'程式名稱
		LAYER_UP		= request.Form ("LAYER_UP")		'上層代碼
		LAYER			= request.Form ("LAYER")		'層級
		VIRTUAL_PATH	= request.Form ("VIRTUAL_PATH")	'虛擬路徑
		

		dim BolFlag,StrSql_Q,ArrVal
		strSql_Q = "SELECT  PROGRAM_ID FROM SYSPROGRAM WHERE PROGRAM_ID='" & trim(PROGRAM_ID) & "'"
		'驗証key是否重覆
		BolFlag	=QueryFun(StrSql_Q,ArrVal)			 	 
		IF BolFlag=true then
			call errortoback("注意!!程式代碼重覆!!!")
		else
		'寫入YSBMSTRE table
			strSql = "INSERT INTO SYSPROGRAM(PROGRAM_ID,PROGRAM_NAME,LAYER_UP,LAYER,VIRTUAL_PATH) VALUES('" & trim(PROGRAM_ID) & "',N'" & trim(PROGRAM_NAME) & "','" & trim(LAYER_UP) & "','"& LAYER &"','"& trim(VIRTUAL_PATH) &"')"
			conn.Execute(strSql)
			conn.Close ()
			set conn=nothing
		end if
	case 1 'UpDate & Delete
		dim ChkDelete,Cnt,iRows
		dim txtPROGRAM_ID,txtPROGRAM_NAME,txtLAYER_UP,txtLAYER,txtVIRTUAL_PATH
		Cnt = Request.Form ("Cnt")	'記載修改資料數量
		
		'one by one 呼叫SP_YSBAE0101_01來update or delete
		for iRows = 0 to Cnt
			ChkDelete		= Request.Form ("ChkDelete"			& iRows)	'刪除checkbox
			txtPROGRAM_ID	= Request.Form ("txtPROGRAM_ID"		& iRows)	'程式代碼
			txtPROGRAM_NAME	= Request.Form ("txtPROGRAM_NAME"	& iRows)	'程式名稱
			txtLAYER_UP		= Request.Form ("TxtLAYER_UP"		& iRows)	'上層代碼
			txtLAYER		= Request.Form ("txtLAYER"			& iRows)	'層級
			txtVIRTUAL_PATH	= Request.Form ("txtVIRTUAL_PATH"	& iRows)	'虛擬路徑

			StrSql="SP_YSBAE0101_01 '"& trim(ChkDelete) &"','"& trim(txtPROGRAM_ID) &"',N'"& trim(txtPROGRAM_NAME) &"','"& trim(txtLAYER_UP) &"','"& trim(txtLAYER) &"','"& trim(txtVIRTUAL_PATH) &"'"
			'Response.Write strsql&"<br>"
			conn.Execute(StrSql)
		next
		conn.Close ()
		set conn=nothing
end select
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
	<INPUT type="hidden" value="<%=PROGRAM_ID%>" id=SearchKey name=SearchKey>
</form>
</body>
</html>