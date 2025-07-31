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
<meta http-equiv="x-ua-compatible" content="IE=10" >
<link rel="stylesheet" type="text/css" href="template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="template/css/mis.css">
</head>

<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0">	
<div id="mainpage">
	<div class="container">
        <div class="col-md-12 mt-2 mb-2" >
            <span class="topheader text-danger"><i class="fa fa-users"></i> 人事管理系統</span>
            <span class="bottomheader text-secondary" >Hệ Thống Quản Lý Nhân Sự</span>
        </div>
    	<!-- Group DIV -->
        <!----------Tạo Dữ Liệu Cơ Bản------------------------------------->
    	<div class="row pt-3 pr-5 pl-5">
             <div class="col-md-4">
            	<a href="javascript:dataclick('1')" class="btn btn-danger btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-folder-open" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">A.基本資料建檔護</font><br>
                            <font class="text-white-70" >Tạo Dữ Liệu Cơ Bản</font>
                        </div>
                    </div>
                </a>
            </div>
            <!-----DL Cơ Bản của NV-------------------------->
            <div class="col-md-4">
            	<a href="javascript:dataclick('2')" class="btn btn-primary btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-address-book" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">B.員工基本資料</font><br>
                            <font class="text-white-70" >DL Cơ Bản của NV</font>
                        </div>
                    </div>
                </a>
            </div>
            <!---------QL Tiền Lương---------------------->
            <div class="col-md-4">
            	<a href="javascript:dataclick('7')" class="btn btn-success btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-money" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">C.薪資管理作業</font><br>
                            <font class="text-white-70" >Quản Lý Tiền Lương</font>
                        </div>
                    </div>
                </a>
            </div>
            <!------------------------------->
        </div>
        <!----------------------------------->
        <!-- Group DIV -->
        <!------Chấm Công NV-------------------->
        <div class="row pt-4 pr-5 pl-5">
             <div class="col-md-4">
            	<a href="javascript:dataclick(3)" class="btn btn-info btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-clock-o" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">D.員工差勤作業</font><br>
                            <font class="text-white-70" >Chấm Công NV</font>
                        </div>
                    </div>
                </a>
            </div>
            <!-----------DL NV Nghỉ Phép-------------------->
            <div class="col-md-4">
            	<a href="javascript:dataclick(4)" class="btn btn-warning text-white btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-coffee" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">E.員工請假資料</font><br>
                            <font class="text-white-70" >DL NV Nghỉ Phép</font>
                        </div>
                    </div>
                </a>
            </div>
            <!-----------QL Báo Biểu DL-------------------->
            <div class="col-md-4">
            	<a href="javascript:dataclick(6)" class="btn btn-puple btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-file-text-o" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">F.報表資料維護</font><br>
                            <font class="text-white-70" >QL Báo Biểu DL</font>
                        </div>
                    </div>
                </a>
            </div>
            <!------------------------------->
        </div>
        <!-- Group DIV -->
        <!---------------DL Máy Gạt Thẻ------------->
        <div class="row pt-4 pr-5 pl-5">
             <div class="col-md-4">
            	<a href="javascript:dataclick(5)" class="btn btn-brown btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-fax" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">G.接收卡鐘資料</font><br>
                            <font class="text-white-70" >DL Máy Gạt Thẻ</font>
                        </div>
                    </div>
                </a>
            </div>
            <!-------Khấu Trừ Lương NV------------------------>
            <div class="col-md-4">
            	<a href="javascript:dataclick(8)" class="btn btn-secondary btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-minus-square" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">H.員工扣款作業</font><br>
                            <font class="text-white-70" >Khấu Trừ Lương NV</font>
                        </div>
                    </div>
                </a>
            </div>
            <!------------Năng Suất NV------------------->
            <div class="col-md-4">
            	<a href="javascript:dataclick(9)" class="btn btn-prink btn-block shadow">
                	<div class="row p-2" >
                    	<div class="col-md-2 pt-3"><span class="fa fa-thumbs-o-up" style="font-size:2rem;"></span></div>
                        <div class="col-md-10" style="text-align:left;">
                        	<font style="font-size:1.2rem;">I.員工考核作業</font><br>
                            <font class="text-white-70" >Năng Suất NV</font>
                        </div>
                    </div>
                </a>
            </div>
            <!------------------------------->
        </div>
    </div>
</div>

</body>
</html>

<script type="text/javascript" language="javascript">
    function dataclick(a) {
        //alert(a);
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
    }
</script>
