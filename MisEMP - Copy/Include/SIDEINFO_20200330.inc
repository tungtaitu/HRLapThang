<!-------- #include file = "../GetSQLServerConnection.fun" --------->
<%

if request("pgid")="" then pgid=session("pgid") else pgid = session("pgid") 
dd1=year(date())&"/"&right("00"&month(date()),2)&"/01"
dd2=year(date())&"/"&right("00"&month(date()),2)&"/"&right("00"&day(date()),2) 
if instr(session("vnlogIP"),"168")>0 then 
	wx="LA" 
elseif	instr(session("vnlogIP"),"169")>0 then  
	wx="DN" 
elseif	instr(session("vnlogIP"),"171")>0 then  
	wx="BC" 
elseif	instr(session("vnlogIP"),"170")>0 then  
	wx="BD" 	
end if 

function formatdat(d) 
	response.write year(d)&"/"&right("00"&month(d),2)&"/"&right("00"&day(d),2) 
end function  
If Session("netuser") = "" Then
    'Response.write "請重新登入!!"
    'response.end
    'response.redirect "../../Login.asp"
%>
	<script language=vbscript>
		alert "UserID is Null Please Login again !!!" & chr(13) &_
			  "Ma so khong ton tai, Xin Hay Dang Nhap Lai He Thong !!! "
		open "../default.asp" , "_top"
		'parent.best.rows="100%,0%"
	</script>
<%
End If
REMOTE_IP = Request.ServerVariables("REMOTE_ADDR")   
session("vnlogIP")=REMOTE_IP 

Dim Str, lenstr, insrev, sname, mname
str= Request.ServerVariables("PATH_INFO")	'http 格式檔名
lenstr=Len(str)								'檔名總字串長度
insrev= InstrRev(str,"/")					'從頭至最後一個 /  的字串長度
sname=right(str,lenstr-insrev)				'含 .asp
mname=mid(sname,1,len(sname)-4)   
show_str ="["& replace(str,"/MisEmp/","") &"]" 
%>
<DIV STYLE="LEFT: 475px;
			MARGIN: 0px;
			POSITION: absolute;
			TOP: 5px;
			WIDTH: 200px; class=txt8"
>
<font class="txt8"><%=show_str%><br>
User : <%=Session("netuser")%>(<%=REMOTE_IP%>) </font>
</DIV>

<%

Set conn = GetSQLServerConnection()
sql="select * from SYSPROGRAM where program_id='"& pgid &"' "
'response.write sql &"<BR>"
set rs=conn.execute(Sql)
if not rs.eof then 
	pgname = rs("program_name")&rs("proname_vn")
else
	pgname=""
end if 
set rs=nothing 
%> 

<DIV STYLE="LEFT: 5px;
			MARGIN: 0px;
			POSITION: absolute;
			TOP: 5px;
			WIDTH: 630px"> 
 
<img src="../image/icon.gif" border=0 align="absmiddle">&nbsp;<font style='font-size:12pt;FONT-FAMILY:Verdana'><%=pgname%></font>
<br>
<hr size=0	style='border: 1px dotted #999999;' align=left  >
</DIV>
 
<script language=vbs>
 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 	 
	if window.event.keyCode = 113  then 
		go()
	end if 	
end function
</script>