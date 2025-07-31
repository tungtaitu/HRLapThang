<!-- #include file = "../GetSQLServerConnection.fun" -->

<link rel="stylesheet" type="text/css" href="../template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="../template/font-awesome/css/font-awesome.css">
<link rel="stylesheet" type="text/css" href="../template/css/mis.css">
<link rel="stylesheet" type="text/css" href="../template/datepicker/datepicker.css">
<link rel="stylesheet" type="text/css" href="../Include/MisStyles.css" >
<%
'Response.write("a="&request("pgid"))
if request("pgid")="" then pgid=session("pgid") else pgid = request("pgid") 

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
	<script language=javascript>
		alert("UserID is Null Please Login again !!!" + chr(13) +  "Ma so khong ton tai, Xin Hay Dang Nhap Lai He Thong !!! ");
		open("../default.asp" , "_top");
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


<%

Set conn = GetSQLServerConnection()
sql="select * from SYSPROGRAM where program_id='"& pgid &"' "
'response.write sql &"<BR>"
set rs=conn.execute(Sql)
if not rs.eof then 
	pgname = rs("program_name")
	pgname_v = rs("proname_vn")
else
	pgname=""
	pgname_v=""
end if 
set rs=nothing 
%> 

<DIV class="col-md-12" STYLE="LEFT: 0px;MARGIN: 0px;POSITION: absolute;TOP: 0px;WIDTH: 98%;"> 
	<table width="100%" border=0 >
		<tr>
			<td style="white-space: nowrap;">
				<span class="topheader text-danger" ><i class="fa fa-folder-open"></i><%=pgname%></span><br>
				<span class="bottomheader text-secondary"><%=UCASE(pgname_v)%></span>
			</td>
			<td align="right" valign="top"><span ><%=show_str%></span></td>
		</tr>
	</table>
</DIV>
  
 