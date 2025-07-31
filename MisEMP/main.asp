<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "GetSQLServerConnection.fun" --> 
<!-- #include file="ADOINC.inc" -->
<%
'response.write "xxx"& session("netuser")
if session("netuser")="" then 
	'response.redirect  "default.asp"
	response.write "<script>"
	response.write "window.open('default.asp','_top');"
	response.write "</script>"
end if  
Set CONN = GetSQLServerConnection()
Set rds = Server.CreateObject("ADODB.Recordset")      
Set rs2 = Server.CreateObject("ADODB.Recordset")      

sqln="select  a.* , b.empnam_cn  from "&_
	 "( select   empid, max(edat) edat  from  empvisadata  group by  empid  ) a   "&_
	 "join  ( select empid , empnam_cn, empnam_vn , outdat from view_empfile ) b  on b.empid = a.empid  "&_
	 "where  isnull(b.outdat,'')=''    "&_
	 "and  convert(char(10),a.edat,111) < convert(char(10), dateadd( d, 20, getdate()) , 111)  "&_
	 "order by   convert(char(10),(edat),111) " 
rds.open sqln , conn, 3, 3   
'response.write sqln 
sqld="select  a.* , b.empnam_cn  from "&_
	 "( select   empid, max(edat) edat  from  empHTdata  group by  empid  ) a   "&_
	 "join  ( select empid , empnam_cn, empnam_vn , outdat from view_empfile ) b  on b.empid = a.empid  "&_
	 "where  isnull(b.outdat,'')=''    "&_
	 "and  convert(char(10),a.edat,111) < convert(char(10), dateadd( d, 20, getdate()) , 111)  "&_
	 "order by   convert(char(10),(edat),111) "  
rs2.open sqld , conn, 3, 3   
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" type="text/css" href="Template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="Template/font-awesome/css/font-awesome.css">
<link rel="stylesheet" type="text/css" href="Template/css/mis.css">
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0">
	
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">	
	<tr>
		<td style="height:20px">
			<div class="col-md-12 mt-2 mb-2" >
				<span class="topheader text-danger"><i class="fa fa-users"></i>人事薪資系統</span>
				<span class="bottomheader text-secondary" >Hệ Thống Quản Lý Nhân Sự</span>
			</div>
		</td>
	</tr>
	<tr>
		<td height="100%" align="center">	
			<table cellspacing="3" cellpadding="3" border=0 style="width:75%">	
				<tr>			
					<td style="width:25%">				
						<a href="javascript:dataclick('1')" class="btn btn-danger btn-block shadow">                	
							<div class="col-md-2 pt-3"><span class="fa fa-folder-open" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">A.基本資料建檔</font><br>
								<font class="text-white-70" >Tạo Dữ Liệu Cơ Bản</font>
							</div>                    
						</a>
					</td>
					<td style="width:25%"> 
						<a href="javascript:dataclick('2')" class="btn btn-primary btn-block shadow">                	
							<div class="col-md-2 pt-3"><span class="fa fa-address-book" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">B.員工基本資料</font><br>
								<font class="text-white-70" >DL Cơ Bản của NV</font>
							</div>                    
						</a>
					</td>
					<td style="width:25%"> 
						<a href="javascript:dataclick('7')" class="btn btn-success btn-block shadow">                	
							<div class="col-md-2 pt-3"><span class="fa fa-money" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">C.薪資管理作業</font><br>
								<font class="text-white-70" >Quản Lý Tiền Lương</font>
							</div>                    
						</a>
					</td>
				</tr> 
				<tr>			
					<td> 
						<a href="javascript:dataclick(3)" class="btn btn-info btn-block shadow">
							<div class="col-md-2 pt-3"><span class="fa fa-clock-o" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">D.員工差勤作業</font><br>
								<font class="text-white-70" >Chấm Công NV</font>
							</div>
						</a>
					</td>		
					<td> 
						<a href="javascript:dataclick(4)" class="btn btn-warning text-white btn-block shadow">
							<div class="col-md-2 pt-3"><span class="fa fa-coffee" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">E.員工請假作業</font><br>
								<font class="text-white-70" >DL NV Nghỉ Phép</font>
							</div>
						</a>
					</td>
					<td> 
						<a href="javascript:dataclick(6)" class="btn btn-puple btn-block shadow">
							<div class="col-md-2 pt-3"><span class="fa fa-file-text-o" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">F.報表資料管理</font><br>
								<font class="text-white-70" >QL Báo Biểu DL</font>
							</div>
						</a>
					</td>
				</tr>	
				<tr>
					<td> 
						<a href="javascript:dataclick(5)" class="btn btn-puple btn-block shadow">
							<div class="col-md-2 pt-3"><span class="fa fa-fax" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">G.接收卡鐘資料</font><br>
								<font class="text-white-70" >DL Máy Gạt Thẻ</font>
							</div>
						</a>				
					</td>
					<td> 
						<a href="javascript:dataclick(8)" class="btn btn-secondary btn-block shadow">
							<div class="col-md-2 pt-3"><span class="fa fa-minus-square" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">H.員工扣款作業</font><br>
								<font class="text-white-70" >Khấu Trừ Lương NV</font>
							</div>
						</a>
					</td>		
					<td> 
						<a href="javascript:dataclick(9)" class="btn btn-prink btn-block shadow">
							<div class="col-md-2 pt-3"><span class="fa fa-thumbs-o-up" style="font-size:2rem;"></span></div>
							<div class="col-md-10" style="height:80px;text-align:center;">
								<font style="font-size:1.2rem;">I.員工績效考核</font><br>
								<font class="text-white-70" >Năng Suất NV</font>
							</div>
						</a>
					</td>
				</tr>		
			</table>
		</td>
	</tr>
</table>

</body>
</html>

<script type="text/javascript" language="javascript">
	
    function dataclick(a) {
        
        if (a == "1") {        
            window.open("function.asp?program_id=A", "contents");
        } else if (a == "2") {      
            window.open("function.asp?program_id=B", "contents");
		} else if (a == 3) {
		    open("function.asp?program_id=D", "contents");
		} else if (a == 4) {
		    open("function.asp?program_id=E", "contents");
		} else if (a == 5) {
		    open("function.asp?program_id=G", "contents");
		} else if (a == 6) {
		    open("function.asp?program_id=F", "contents");
		} else if (a == 7) {
		    open("function.asp?program_id=C", "contents");
		} else if (a == 8) {
		    open("function.asp?program_id=H", "contents");
		} else if (a == 9) {
		    open("function.asp?program_id=I", "contents");
		}
		
		var tdLeftMenu=window.parent.document.getElementById('tdLeftMenu');
		var btn_MenuShow=window.parent.document.getElementById('btn_MenuShow'); 
		var btn_MenuHide=window.parent.document.getElementById('btn_MenuHide');
		
		
		tdLeftMenu.style.display = "block";
		btn_MenuShow.style.display = "none";
		btn_MenuHide.style.display = "block";
		
    }
</script>

<script language="VBScript">
	function dataclick1(a)
        
		if a = 1 then			
			open "function.asp?program_id=A" , "contents"			
		elseif a = 2 then
			open "function.asp?program_id=B" , "contents"
		elseif a = 3 then			
			open "function.asp?program_id=D" , "contents"
		elseif a = 4 then
			open "function.asp?program_id=E" , "contents"
		elseif a = 5 then
			open "function.asp?program_id=G" , "contents"
		elseif a = 6 then
			open "function.asp?program_id=F" , "contents"
		elseif a = 7 then
			open "function.asp?program_id=C" , "contents"
		elseif a = 8 then
			open "function.asp?program_id=H" , "contents"
		elseif a = 9 then
			open "function.asp?program_id=I" , "contents"
		end if
	end function
</script>