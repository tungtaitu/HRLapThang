<%@Language=VBScript Codepage=65001 %>
<%Response.Buffer =True%>  
<!--#include file="../GetSQLServerConnection.fun"-->  
<!-- #include file="../Include/global_asp_fun.asp" -->
<!-- #include file="../Include/global_vbs_fun.asp" -->
<%
const self	="YSBAE0401.FORE.ASP"
const action	="YSBAE0401.UPDATEDB.ASP"

const formname	="FRM"
const method	="POST"  

%>
<%
dim conn,iRows
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open GetSQLServerConnection()


'查詢條件值
dim SearchKey
SearchKey=request.Form ("SearchKey")
if SearchKey="" then SearchKey="*"


'取群組代碼
dim StrSql_GROUP,BolFlag_GROUP,Arrdata_GROUP
StrSql_GROUP="select sys_type, sys_value  from BasicCode  where func = 'Grp' order by sys_type"
'response.write StrSql_GROUP
'response.end 
BolFlag_GROUP	=QueryFun(StrSql_GROUP,Arrdata_GROUP)

'查詢群組
dim GROUP_ID
GROUP_ID=request.Form ("GROUP_ID")
if GROUP_ID="" then GROUP_ID=trim(Arrdata_GROUP(0,0))





'分頁清單列示
dim StrSql
select case trim(SearchKey)
	case "*"
		StrSql="SP_YSBAE0401_01 '"& trim(GROUP_ID) &"',''"
	case else
		StrSql="SP_YSBAE0401_01 '"& trim(GROUP_ID) &"','"& trim(SearchKey) &"'"
end select 

dim size,flag,page,pagecount,recordcount,Getdata
size=300
page=Request.Form ("page")
if isnumeric(page)=false then page=1
'------------call abspage(分頁function)
if StrSql <> empty then
	flag=AbsPage(size,StrSql,page,pagecount,recordcount,Getdata)
else
	flag		=false
end if

if flag=false then
	page		=0
	pagecount	=0
	recordcount	=0
end if

conn.Close ()
set conn=nothing
%>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=UTF-8">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>

<script language =vbs>
'初始化function
function Start()
	'READONLY
	'call diffcolor("CUSTID",1)
	'call diffcolor("CUSTSNAME",1)

end function

'傳入值編審 for 新增
function chkval()
	dim Elm
	for each Elm in <%=formname%>
		select case Elm.name
			case "COMNO" '驗証新增之統編
				dim nErrCode,sErrDesp
				if CheckCompanyId(Elm.value,nErrCode,sErrDesp)=false then
					alert sErrDesp
					chkval=false
					Elm.focus
					exit function
				end if
		end select
	next
	chkval=true
end function

'傳入值編審 for 修改
'function chkval_update()
'	dim Elm
'	for each Elm in <%=formname%>
'		'驗証修改之統編
'		select case elm.name
'			case "comno"
'				if CheckCompanyId(Elm.value,nErrCode,sErrDesp)=false then
'					alert sErrDesp
'					chkval_update=false
'					Elm.focus
'					exit function
'				end if
'		end select
'	next
'	chkval_update=true
'end function


function go(index)
	with window.<%=formname%> 
		select case index
			case 0 'AddNew
				if chkval()=true then
					call UcaseChar()
					.action="<%=action%>"
					.ActStatus.value="0"
					if dblconfirm("新增")=true then .submit
				end if
			case 1 'Search
				.page.value=1
				.action="<%=self%>"
				.submit
			case 2 'Update & Delete
				'if chkval_update()=true then
					call UcaseChar()
					.action="<%=action%>"
					.ActStatus.value="1"
					if dblconfirm("異動")=true then .submit
				'end if
			case 3 'Page first
				.page.value=1
				.action="<%=self%>"
				.submit
			case 4 'Page prev
				.page.value=cint(.page.value)-1	
				.action="<%=self%>"
				.submit
			case 5 'Page next
				.page.value=.page.value+1
				.action="<%=self%>"
				.submit
			case 6 'Page last
				.page.value=<%=pagecount%>
				.action="<%=self%>"
				.submit
		end select
			
	end with	
end function

</script>
<body  scroll='auto' onkeydown="enterto()" topmargin=5 >
<form NAME="<%=formname%>" ACTION="<%=action%>" method="<%=method%>"   >
<table width="460" border="0" cellspacing="0" cellpadding="0">
  <tr>
   	<TD ><img border="0" src="../image/icon.gif" align="absmiddle">
   	<%=session("pgname")%></TD>		 
  </tr>
</table> 	 
<hr size=0	style='border: 1px dotted #999999;' align=left width=550>	
<table border="0" class=txt   width=550>
	<tr>
		<td width =100%>
			請選擇群組
			<SELECT ID="GROUP_ID" NAME="GROUP_ID" style="width:140" class=txt onchange="go(1)">
			<%for iRows=0 to ubound(Arrdata_GROUP)%>
			<OPTION value =<%=Arrdata_GROUP(iRows,0)%> <%if Arrdata_GROUP(iRows,0)=GROUP_ID then%> selected <%end if%>><%=Arrdata_GROUP(iRows,0)%>-<%=Arrdata_GROUP(iRows,1)%></OPTION>
			<%next%>
			</SELECT>&nbsp;&nbsp;
			請輸入程式代碼或名稱&nbsp;&nbsp;<input type="text" class="input4" name="SearchKey" id="Text3" size="15" maxlength=80 value="<%=SEARCHKEY%>">
			<INPUT type="button" value="查  詢" id="BtnSearch" name=BtnSearch class="button" onclick="go(1)" onkeydown="go(1)" >
			<INPUT type="hidden" value="<%=Page%>" id=page name=page>
			<INPUT type="hidden" value="" name=ActStatus ID="ActStatus">
		</td>
	</tr>
	<tr>
		<td width=100%>
		<hr size=0	style='border: 1px dotted #999999;' align=left >	 
		</td>
	</tr>
</table>
<table border="0" class=txt  width=550>
	<%if flag=true then%>
	<tr BGCOLOR=#B4C5DA>
		<td width=60>Pro_ID</td>
		<td width=180>Pro_Name</td>
		<td width=180>Pro_Name(VN)</td>
		<td width=65 align=center>Read</td>
		<td width=65 align=center >Edit</td>
	</tr>
	<%for iRows=0 to ubound(Getdata)
	%>
	<tr>
		<td >
			<input type="text" class="readonly2" name="Program_id<%=iRows%>" id="Program_id<%=iRows%>" size="10" maxlength=20 value="<%=getdata(iRows,0)%>" readonly  >
		</td>
		<td>	
			<input type="text" class="readonly2" name="Program_name<%=iRows%>" id="Program_name<%=iRows%>" size="27" maxlength=20 value="<%=getdata(iRows,1)%>" readonly >
		</td>
		<td>	
			<input type="text" class="readonly2" name="Program_name<%=iRows%>" id="Program_name<%=iRows%>" size="27" maxlength=20 value="<%=getdata(iRows,4)%>" readonly  >
		</td>
		<td align=center>
			<INPUT type="checkbox" ID="GROUP_R<%=iRows%>" NAME="GROUP_R<%=iRows%>" <%if trim(getdata(iRows,2))="Y" then%> checked <%end if%>>
		</td>
		<td align=center>
			<INPUT type="checkbox" ID="GROUP_W<%=iRows%>" NAME="GROUP_W<%=iRows%>" <%if trim(getdata(iRows,3))="Y" then%> checked <%end if%>>
		</td> 
	</tr>
	<%	next
		else%>
	<tr>
		<td width =100%>目前無相關資料!!!</td>
	</tr>
	<%	end if%>
	
	
</table>
<%if flag=true then%>
<table border="0" class=txt   width=550>
	<tr>
		<td width=100%><hr></td>
	
	</tr>
	<td width="100%" align="center" >
	頁次:<%=Page%>/<%=PageCount%>　總筆數:<%=recordcount%>
	</td>
	<tr>
		<td width =100%  align =center >
			<INPUT type="button" value="第一頁" name="btn_first" class=button onclick="go(3)" ID="btn_first">
			<INPUT type="button" value="上一頁" name="btn_prev" class=button onclick="go(4)" ID="btn_prev">
			<INPUT type="button" value="下一頁" name="btn_next" class=button onclick="go(5)" ID="btn_next">
			<INPUT type="button" value="最末頁" name="btn_last" class=button onclick="go(6)" ID="btn_last">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT type="hidden" value="<%=iRows-1%>" id="cnt" name=cnt>
			<%if UCASE(session("mode"))="W" then%>
				<INPUT type="button" value="確  定" id=BtnUpdate name=BtnSure class=button onclick="go(2)">
				<INPUT type="reset"  value="取  消" id=BtnRst  name=BtnRst class=button>
			<%end if%>
		</td>
	</tr>
</table>
<%end if%>
</form>
</body>
</html>
