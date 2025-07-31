<!-- #include file="../../Include/sideinfo.inc" -->
<!--#include file="../../include/ADOINC.inc"-->
<!--#include file="../../GetSQLServerConnection.fun"-->
<!-- #include file="../../Include/css.inc" -->
<%
dim self,action,formname
self   ="YSBAE0401.ASP"
action ="YSBAE0401.UPDATEDB.ASP"
formname="frm_update"

%>
<%'--程式權限判斷%>
<!--#include file="../../include/checkpower.asp"-->
<%
on error resume next
'群組代碼
 dim group_id
 group_id=Request("group_id")
 session("group_id")=group_id
 'Response.Write session("group_id")


dim conn 'conn變數
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open GetSQLServerConnection()

dim sql_dropdown 
sql_dropdown="proc_sysgroup_dropdown" '下拉式選單 sp

dim rs_dropdown
set rs_dropdown = server.CreateObject("ADODB.Recordset")
rs_dropdown.CursorLocation = 3
rs_dropdown.Open sql_dropdown,conn,3,1

'查詢條件值
dim search_key
search_key=trim(request("search_key"))



dim size
size=10
'送值進store proc
dim Source
Source = "proc_syspower "&size&",'"&session("group_id")&"','"&search_key&"'"
'--- 設定一個 Recordset
'分頁筆數 


Dim objList	'recordset 變數
set objList = server.CreateObject("ADODB.Recordset")
objList.CursorLocation = 3
objList.open Source,conn,3,1
objList.PageSize = size

'--- 分頁設定
Dim intPage '頁數
Dim intUp	'向上一頁
Dim intDown	'向下一頁
Dim intPageCount '頁數

intPage = Request("page")
intPageCount = objList.PageCount

if intPage <= 0 or intPage = "" then
	intPage = 1
else 
	intPage = int(request("page"))
	if intPage > intPageCount then
		intPage = intPageCount
	end if
end if

if intPage-1 <= 0 then
	intUp = 1
else
	intUp = intPage -1
end if

if intPage + 1 >= intPageCount then
	intDown = intPageCount
else
	intDown = intPage + 1
end if

objList.AbsolutePage = intPage 

dim recordcount 
recordcount=objList("RECORDCOUNT")

%>

<html>
<head>
<link rel="stylesheet" href="./../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<!--#include file="../../include/global_vbs_fun.asp"-->
<script language="vbs">

function sta()
  
  dim status
  status="<%=request("status")%>"
  if status<>"" then 
    alert ("群組名稱"&status&"重復!")
  end if
end function

function go(byval page)
  window.frm_update.page.value = page
  call UcaseChar()
  window.frm_update.submit  
end function 

function change(program_sn)
  window.frm_update.change_sn.value =trim(cstr(window.frm_update.change_sn.value)) +"'"+trim(cstr(program_sn))+"',"
end function

function change_group()
  window.frm_search.submit  
end function


</script>
</head>

<body onload="sta()" background="..\..\Picture\bg9.gif">
<table width=630 class=txt12><tr><td align=center><b><%=session("pgname")%></b></td></tr></table>

<table width=630><tr><td align=center> 
	<table width=547 border=0  cellpadding=0 cellspacing=0 class=txt>
		<tr class=txt>
			<td  ></td>	
			<td  >	
			<FORM NAME =frm_search ACTION="<%=self%>" method="get" >
				<table width=500 class=txt><tr>						
				<td>群組名稱：</td>
				<td>
				<SELECT id=group_id name=group_id onchange="change_group()">				                        <OPTION></OPTION>
				                        <%do
				                          if rs_dropdown.eof=true then exit do %>				                        <OPTION value="<%=rs_dropdown("group_id")%>" <%if trim(rs_dropdown("group_id"))=trim(session("group_id")) then%> selected <%end if%>><%=rs_dropdown("group_name")%></OPTION>				                        <%rs_dropdown.MoveNext 				                          loop%>
				</SELECT></td>
				<td><font class=font1>程式名稱：</font></td>			
				<td><INPUT type="text" id=search_key name=search_key class=input1></td>
				<td><INPUT type="submit" value="查  詢" id=submit1 name=submit1 class=button <%if session("group_id")="" then%> disabled <%end if%>></td>
				</tr>
				</table>	
			</form>			
			<FORM NAME =frm_update ACTION="<%=action%>?pgid=<%=session("pgid")%>&pgname=<%=session("pgname")%>" method="post">
                <table border="0"  width="500" align="center" height="100" CLASS=TXT> 
                    <tr bgcolor=LightGrey >
                        <td width="130" align="center" >程式代號</td>
                        <td width="230" align="center" >程式名稱</td>
                        <td width="70" align="center"  >讀取權限</td>
                        <td width="70" align="center"  >存取權限</td>
                        
                        
                    </tr>
                        <%for i = 1 to objList.PageSize%>
							 <%if i mod 2 = 1 then %>                      
                    <tr  >
                        <INPUT type="hidden" id=PROGRAM_SN<%=i%> name=PROGRAM_SN<%=i%> VALUE="<%=objList("PROGRAM_SN")%>" >  
                        <td align="center"  ><INPUT type="text" class=inputBOX id=PROGRAM_ID<%=i%> name="PROGRAM_ID<%=i%>" size=17 VALUE="<%=objList("PROGRAM_ID")%>"  maxlength=30  disabled ></td>
                        <td align="center"  ><INPUT type="text" class=inputbox id=PROGRAM_NAME<%=i%> name="PROGRAM_NAME<%=i%>" size=34 VALUE="<%=objList("PROGRAM_NAME")%>"  maxlength=30 disabled ></td>
                        <td align="center"  ><INPUT TYPE="checkbox" border=0 id=checkbox_r<%=i%> name="checkbox_r<%=i%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if ((UCASE(session("mode"))="R" or objlist("rw")="0") or session("group_id")="") then%> disabled <%end if%> onchange="change(<%=objList("PROGRAM_SN")%>)" <%if trim(objList("r"))="1" then%> checked <%end if%>></td>                         
                        <td align="center"  ><INPUT TYPE="checkbox" border=0 id=checkbox_w<%=i%> name="checkbox_w<%=i%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if ((UCASE(session("mode"))="R" or objlist("rw")="0") or session("group_id")="") then%> disabled <%end if%> onchange="change(<%=objList("PROGRAM_SN")%>)" <%if trim(objList("w"))="1" then%> checked <%end if%>></td>                         
                    </tr>
                             <%else%>
                    <tr  >
                        <INPUT type="hidden" id=PROGRAM_SN<%=i%> name=PROGRAM_SN<%=i%> VALUE="<%=objList("PROGRAM_SN")%>" >  
                        <td align="center" ><INPUT type="text" class=inputbox id=PROGRAM_ID<%=i%> name="PROGRAM_ID<%=i%>" size=17 VALUE="<%=objList("PROGRAM_ID")%>" maxlength=30 disabled ></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=PROGRAM_NAME<%=i%> name=PROGRAM_NAME<%=i%> size=34 VALUE="<%=objList("PROGRAM_NAME")%>"  maxlength=30 disabled ></td>
                        <td align="center" ><INPUT TYPE="checkbox" border=0 id=checkbox_r<%=i%> name="checkbox_r<%=i%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if ((UCASE(session("mode"))="R" or objlist("rw")="0") or session("group_id")="") then%> disabled <%end if%> onchange="change(<%=objList("PROGRAM_SN")%>)" <%if trim(objList("r"))="1" then%> checked <%end if%>></td>                         
                        <td align="center" ><INPUT TYPE="checkbox" border=0 id=checkbox_w<%=i%> name="checkbox_w<%=i%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if ((UCASE(session("mode"))="R" or objlist("rw")="0") or session("group_id")="") then%> disabled <%end if%> onchange="change(<%=objList("PROGRAM_SN")%>)" <%if trim(objList("w"))="1" then%> checked <%end if%>></td>                         
                    </tr>
                             <%end if%>     
                             
                        <%	
							objList.movenext
							if objList.EOF then exit for
                        Next%>
                    <tr>
                        <td width="50%" align="center" colspan="6" height="25" valign="middle" class=td1>
                            <INPUT type="button" value="第一頁" name="button4" class=button onclick="go(1)">
                            <INPUT type="button" value="上一頁" name="button1" class=button onclick="go(<%=intUp%>)">
                            <INPUT type="button" value="下一頁" name="button2" class=button onclick="go(<%=intDown%>)">
                            <INPUT type="button" value="最末頁" name="button5" class=button onclick="go(<%=intPageCount%>)">
                            <INPUT type="hidden" value="<%=intPage%>" id=page name=page>
                            <INPUT type="hidden" value="<%=objList.PageSize%>" id=pagesize name=pagesize>
                            <INPUT type="hidden" value="" id=change_sn name=change_sn>
                            <INPUT type="hidden" value="<%=intPage%>" id=intPage name=intPage >
                            <INPUT type="hidden" value="<%=intPageCount%> " id=intPageCount name=intPageCount>
                            <INPUT type="hidden" value="<%=recordcount%>" id=recordcount name=recordcount>
                            <INPUT type="hidden" value="<%=search_key%>" id=search_key name=search_key>
                            <%if UCASE(session("mode"))="W" then%>
                            <INPUT type="button" value="確  定" id=button3 name=button3 class=button onclick="go(<%=intPage%>)">
                            <INPUT type="reset"  value="取  消" id=reset1 name=reset class=button>
                            <%end if%><BR>
                           目前所在頁數:<%=intPage%>  
                            / 總頁數:<%=intPageCount%>                      
                            / 總筆數：<%=recordcount%>
                         </td>
                    </tr>
                </table>
				</form>
				</td>
			<td></td>
		</tr>
	</table>		
</td></tr></table>
<%set conn=nothing%>
</body>
</html>