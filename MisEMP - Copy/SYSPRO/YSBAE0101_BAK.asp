	<!-- #include file="../../Include/sideinfo.inc" -->
<!--#include file="../../include/ADOINC.inc"-->
<!--#include file="../../GetSQLServerConnection.fun"-->
<!-- #include file="../../Include/css.inc" -->
<%
dim self,action,formname
self   ="YSBAE0101.asp"
action ="YSBAE0101.updatedb.asp"
formname="frm_update"

dim size
size=10
%>
<%'--�{���v���P�_%>
<!--#include file="../../include/checkpower.asp"-->
<%
Response.Write mode2 
%>

<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open GetSQLServerConnection()

'�d�߱����
dim search_key
search_key=trim(request("search_key"))


'�e�ȶistore proc
Source = "proc_sysprogram "&size&",'"&search_key&"'"

Dim objList	'recordset �ܼ�

'--- �]�w�@�� Recordset
set objList = server.CreateObject("ADODB.Recordset")
objList.CursorLocation = 3
objList.open Source,conn,3,1
objList.PageSize = size

'--- �����]�w
Dim intPage '����
Dim intUp	'�V�W�@��
Dim intDown	'�V�U�@��
Dim intPageCount '����

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
    alert ("�{���N��"&status&"���_!")
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

function search(page)
  window.navigate "<%=self%>?page="&page
end function



</script>
</head>
<body onload="sta()" background="..\..\Picture\bg9.gif"  topmargin=0>
<table width=630 class=txt12><tr><td align=center><b><%=session("pgname")%></b></td></tr></table>

<table width=630><tr><td align=center> 
	<table width=547 border=0  cellpadding=0 cellspacing=0 class=txt >
		<tr class=txt>
			<td ></td>	
			<td >							
				<FORM NAME =frm_search ACTION="<%=self%>" method="get">					
				<table border="0" cellpadding="3" cellspacing="3" align="center" class=txt width=500>
                <tr>
                  <td align=right width=135>�{���W��: </td>
                  <td>                    
					<INPUT type="text" id=search_key name=search_key class=inputbox>
					<INPUT type="submit" value="�d  ��" id=submit1 name=submit1 class=button>															
                  </td>                  
                </tr>
              </table> 
              </form>             
              <FORM NAME =frm_update ACTION="<%=action%>" method="post">                            
                <table border="0" cellpadding="1" cellspacing="1" width="500" align="center"   class=txt> 
                    <tr bgcolor=LightGrey >
                        <td width="30" align="center" height="25" >�R��</td>
                        <td width="60" align="center"   >�{���N��</td>
                        <td width="120" align="center" >�{���W��</td>
                        <td width="60" align="center" >�W�h�N��</td>
                        <td width="60" align="center" >�{���h��</td>
                        <td width="170" align="center" >�{�����|</td>
                        
                    </tr>
                        <%for i = 1 to objList.PageSize%>
							 <%if i mod 2 = 1 then %>                      
                    <tr  >
                        <INPUT type="hidden" id=PROGRAM_SN<%=i%> name=PROGRAM_SN<%=i%> VALUE=<%=objList("PROGRAM_SN")%> >  
                        <td align="center" height="25" ><INPUT TYPE="checkbox" border=0 id=checkbox<%=i%> name="checkbox<%=i%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>                         
                        <td align="center" ><INPUT type="text" class=inputbox id=PROGRAM_ID<%=i%> name="PROGRAM_ID<%=i%>" size=7 VALUE="<%=objList("PROGRAM_ID")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=PROGRAM_NAME<%=i%> name="PROGRAM_NAME<%=i%>" size=15 VALUE="<%=objList("PROGRAM_NAME")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=LAYER_UP<%=i%> name=LAYER_UP<%=i%> size=4 VALUE="<%=objList("LAYER_UP")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=LAYER<%=i%> name=LAYER<%=i%> size=4 VALUE="<%=objList("LAYER")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=VIRTUAL_PATH<%=i%> name=VIRTUAL_PATH<%=i%> size=25 value="<%=objList("VIRTUAL_PATH")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        
                    </tr>
                             <%else%>
                    <tr  >
                        <INPUT type="hidden" id=PROGRAM_SN<%=i%> name=PROGRAM_SN<%=i%> VALUE=<%=objList("PROGRAM_SN")%> >  
                        <td align="center" height="25" ><INPUT TYPE="checkbox" border=0 id=checkbox<%=i%> name="checkbox<%=i%>" onchange="change(<%=objList("PROGRAM_SN")%>)"></td>                         
                        <td align="center" ><INPUT type="text" class=inputbox id=PROGRAM_ID<%=i%> name=PROGRAM_ID<%=i%> size=7 VALUE="<%=objList("PROGRAM_ID")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=PROGRAM_NAME<%=i%> name=PROGRAM_NAME<%=i%> size=15 VALUE="<%=objList("PROGRAM_NAME")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=LAYER_UP<%=i%> name=LAYER_UP<%=i%> size=4 VALUE="<%=objList("LAYER_UP")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=LAYER<%=i%> name=LAYER<%=i%> size=4 VALUE="<%=objList("LAYER")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        <td align="center" ><INPUT type="text" class=inputbox id=VIRTUAL_PATH<%=i%> name=VIRTUAL_PATH<%=i%> size=25 value="<%=objList("VIRTUAL_PATH")%>" onchange="change(<%=objList("PROGRAM_SN")%>)" <%if UCASE(session("mode"))="R" then%> disabled <%end if%>></td>
                        
                    </tr>
                             <%end if%>     
                             
                        <%	
							objList.movenext
							if objList.EOF then exit for
                        Next%>
                    </table>                    
                    <table width=500 class=txt>    
                    <tr>
                        <td  align="center"  height="25" valign="middle" >
                            <INPUT type="button" value="�Ĥ@��" name="button4" class=button onclick="go(1)">
                            <INPUT type="button" value="�W�@��" name="button1" class=button onclick="go(<%=intUp%>)">
                            <INPUT type="button" value="�U�@��" name="button2" class=button onclick="go(<%=intDown%>)">
                            <INPUT type="button" value="�̥���" name="button5" class=button onclick="go(<%=intPageCount%>)">
                            <INPUT type="hidden" value="<%=intPage%>" id=page name=page>
                            <INPUT type="hidden" value="<%=objList.PageSize%>" id=pagesize name=pagesize>
                            <INPUT type="hidden" value="" id=change_sn name=change_sn>
                            <INPUT type="hidden" value="<%=search_key%>" id=search_key name=search_key>
                            <%if UCASE(session("mode"))="W" then%>
                            <INPUT type="button" value="�T  �w" id=button3 name=button3 class=button onclick="go(<%=intPage%>)">
                            <INPUT type="reset"  value="��  ��" id=reset1 name=reset1 class=button>
                            <%end if%><br>                        
                             �ثe�Ҧb����:<%=intPage%>  
                            / �`����:<%=intPageCount%>                      
                            / �`���ơG<%=recordcount%>
                         </td>
                    </tr>
                </table>
			</form>
			</td>
			<td background="../../picture/line_right.gif"></td>
		</tr>
	</table>		
</td></tr></table>	

<%set conn=nothing%>

</body>
</html>