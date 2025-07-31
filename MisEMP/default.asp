<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "GetSQLServerConnection.fun" --> 
<!-- #include file="ADOINC.inc" -->
<%
'response.end 
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

code = Ucase(Request("code"))
pwd = Request("pwd")

if request("logout")="Y" then 
	Session.Abandon 	
	response.redirect "default.asp"
end if 
session("rpt")="http://172.22.166.30/MisEmpRpt/"
'session("rpt")="http://localhost/MisEmpRpt/"
'response.write code 
'response.write pwd 
msg=""
If code = "" or pwd="" Then    
	logF="N"
	msg=""
Else
	Set DataSource = GetSQLServerConnection() 
	'Session("DB_Conn") = GetSQLServerConnection() 	
	Set cmd = Server.CreateObject ("ADODB.Command")
	cmd.ActiveConnection = DataSource
	cmd.CommandType = adCMDStoredProc
	cmd.CommandText = "validate"
	cmd.Parameters.Refresh
	cmd("@name") = code
	cmd("@passwd") = pwd	
	cmd("@LastLogin") = "" 
	cmd.Execute
	'response.write "a="&adCMDStoredProc &"<br>"
    'response.write "b="&DataSource &"<br>"
    'response.write "c="&cmd.Execute
    
	If cmd("@LastLogin") = "NO" Then
		msg = "<font color=yellow  >輸入錯誤!!</font>"
		logF="N"
		Session("NETUSER")=""
		SESSION("NETUSERNAME")="" 
		Session("USERPWD") = ""
		Session("LastLogin") = ""
		Session("NETUSERTYPE") = ""
		Session("NETWHSNO") = ""
		Session("NETUNITNO") = ""
		Session("NETF2") = ""
		Session("NETD2") = ""
		Session("NETB2") = ""
		Session("RIGHTS") = "" 		
	Else
		Session("NETUSER") = Ucase(code)
		SESSION("NETUSERNAME") = Ucase(cmd("@uname"))
		Session("USERPWD") = pwd
		Session("LastLogin") = cmd("@LastLogin")
		Session("NETUSERTYPE") = cmd("@loginType")
		Session("NETWHSNO") = cmd("@WHSNO")
		Session("NETUNITNO") = cmd("@UNITNO") 		
		Session("NETG1") = cmd("@B1")
		Session("NETG2") = cmd("@B2")  
		Session("NETF2") = cmd("@F2")
		Session("RIGHTS") = cmd("@RIGHTS")
		DataSource.Close	
		set cmd=nothing 
		set DataSource=nothing
		
		logF="Y"
		
		REMOTE_IP = Request.ServerVariables("REMOTE_ADDR")   
		session("vnlogIP")=REMOTE_IP  
		'Response.Write Session("NETD2") &"<br>"
		msg = "<font color=#FFFFFF>已登入!!</font>"
		Response.redirect "system.asp"		
	End If  	
End If 

Set CONN = GetSQLServerConnection()
Set rds = Server.CreateObject("ADODB.Recordset")      

sqln="select a.empid, b.empnam_cn, convert(char(10),max(edat),111) edat from empvisadata a, ( select empid , empnam_cn, empnam_vn , outdat from empfile )  b   "&_	
	 "where a.empid = b.empid and isnull(b.outdat,'')=''  "&_
	 "group by a.empid , b.empnam_cn having  convert(char(10),max(edat),111)  < convert(char(10), dateadd(  d, 20, getdate()) , 111) order by   convert(char(10),max(edat),111) " 



rds.open sqln , conn, 3, 3 
'response.write "msg=" & msg   

local_ip=trim(Request.ServerVariables("LOCAL_ADDR"))
sqlx="select * from basicCode where func='Xip' and sys_value like '%"& local_ip &"%' "  
'response.write sqlx 
set rsx=conn.execute(Sqlx)
if not rsx.eof then 
	mywhsno=rsx("sys_type")
else
	mywhsno="LT"
end if 	
	
set rsx=nothing 
session("mywhsno") = mywhsno  


 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
	<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="x-ua-compatible" content="IE=10" >
	
	<link rel="stylesheet" href="Include/style.css" type="text/css">
	<link rel="stylesheet" href="Include/style2.css" type="text/css">

	<meta name="viewport" content="width=device-width, initial-scale=1">
	<script src="template/js/jquery-3.4.1.min.js"></script>
	<script src="template/bootstrap/js/bootstrap.min.js"></script>
	<link rel="stylesheet" type="text/css" href="template/bootstrap/css/bootstrap.min.css">
	<link rel="stylesheet" type="text/css" href="template/font-awesome/css/font-awesome.css">
	<link rel="stylesheet" type="text/css" href="template/css/mis.css">

	<title>人事薪資系統</title>
</head>

<style>
body, html {
  height: 100%;
  margin: 0;
}
</style>

<body topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  ONLOAD=f()> 	
<FORM NAME="logon" METHOD=POST action="default.asp">
<div class="divlogin" style="background:url(template/img/peoplebackgroup.jpg) no-repeat top left #FFF;">
    	<div class="row">
            <div class="col-md-8" style="text-align:center; padding:50px;">
                <img src="template/img/logohoaduong2.pngxx" height="57" /><br /><br />
                <font style="font-weight:bolder;font-size:52px; color:red;">人事</font>
                <font style="color:#000; font-size:43px; font-weight:lighter; ">管理系統</font>
                <p style="color:#333; font-size:16px;">員工管理 - 部門管理 - 加班管理 - 薪資管理 - 請假管理 - 保險管理</p>
            </div>
            <div class="col-md-4">
                    <div class="modal-content" style="box-shadow:none; border:none">
                      <div class="modal-header" style="text-align:center; background:#9F0000; color:#FFF;">
                        <h2 style="font-weight:bold" ><span class="fa fa-sign-in"></span> 帳號登入</h2>
                      </div>
                      <div class="lagbox">
                        <a href=""><img src="template/img/vnflag-round-72.png" width="42" /></a>
                      </div>
                      <div class="modal-body">
                            <div class="form-group">
                              <label><span class="fa fa-user"></span> 輸入帳號</label>
							  <input name=CODE id="CODE" size=15 class="form-control" value="" >                              
                            </div>
                            <div class="form-group">
                              <label ><span class="fa fa-lock"></span> 帳號密碼</label>
                              <input type=password name=PWD id="PWD" size=15 class="form-control" value="" maxlength=50>
                            </div>
                            <button type="submit" class="btn btn-default btn-success btn-block">
                            	<span class="glyphicon glyphicon-log-in"></span>  登入</button>
                          
                      </div>
                      <div class="modal-footer" style="background:#ccc;">
                                <%=msg%>&nbsp;&nbsp;
                        </div>
            	</div>
            </div>
        </div>
    </div>
   <nav class="fixed-bottom" style="height:30px;">
   		<div style="display:block; text-align:center; color:#999; margin-top:5px;">
   			<p><i>&copy; M.I.S company</i> | 系統故障請打 : <font color="#FF0000">0909 xxx xxx</font></p>
    	</div>
   </nav>	
</FORM>
</body>

<script language=javaScript>
function f(){
  logon.CODE.focus ();
  logon.CODE.SELECT();
}

function go(){  
	if(logon.CODE.value=="")
	{ 
		alert("input User Name");
	}else if(logon.PWD.value=="")
	{
		alert("input Pwd");
	}else{
		logon.action="default.asp";
		logon.submit();
	}
}

function enterto(){
	if(window.event.keyCode == 13) window.event.keyCode =9 ;	
}

</script> 
