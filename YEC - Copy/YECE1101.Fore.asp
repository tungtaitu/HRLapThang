<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->

<%

self="YECE1101"  
nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="1" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="1" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 	

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

whsno= request("whsno")
if whsno="" then whsno=session("mywhsno")
dfudamt =request("dfudamt")
if dfudamt="" then dfudamt="4400000"

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 20    'number of fields per record 
sortby = request("sortby")
query=trim(request("query"))
if sortby="" then 
	sortstr="isnull(a.mdtm,a.keyindate) desc, len(a.aid) , aid "
elseif sortby="1" then 
	sortstr="a.empid  "
elseif sortby="2" then 
	sortstr="b.gstr , a.empid  "
else 
	sortstr="isnull(a.mdtm,a.keyindate), len(a.aid) , aid "
end if 
qdv = request("qdv")


Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 

sql="select a.* , b.empnam_cn+b.empnam_vn as empname, b.gstr, b.nindat from  "&_
		"(select * from empNoTax where isnull(sts,'')<>'D' ) a "&_
		"join (select * from view_empfile ) b on b.empid = a.empid "&_
		" where 1=1 "
if query<>"" then 
	sql=sql&" and  ( charindex( '"&query&"' , b.empnam_cn+b.empnam_vn )  >0 or a.empid like '%"& query&"%' )"
end if 
if qdv<>"" then  
	sql=sql&" and   b.groupid ='"& qdv &"' "
end if 
sql=sql& 	"order by "& sortstr 'isnull(a.mdtm,a.keyindate), len(a.aid) , aid  
		'response.write sql 
		'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 		
		pagerec= rs.RecordCount +10				
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
		whsno = rs("whsno")
	END IF 	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 
	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("whsno"))
				tmpRec(i, j, 2) = trim(rs("empid"))
				tmpRec(i, j, 3) = trim(rs("empname"))
				tmpRec(i, j, 4) = rs("gstr")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = formatnumber(rs("person_qty"),0)
				tmpRec(i, j, 7) = formatnumber(rs("ut_mtax"),0)				
				tmpRec(i, j, 8) = formatnumber(rs("tot_mtax"),0)				
				tmpRec(i, j, 9) = rs("aid")		
				rs.MoveNext 
			else 
				exit for 
			end if 
		next 	
		 if rs.EOF then 
			rs.Close 
			Set rs = nothing
			exit for 
		 end if 			
		'Session("YECE1101B") = tmpRec			 
	next 	
end if  
set rs=nothing  

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<script type="text/javascript">
	function chknum(sid){
		var code = document.getElementById(sid).value ;
		if (code!="") {
				if ( !isNumeric(code) ) {
				document.getElementById(sid).value ="";
				alert ("nhap so請輸入數字") ;			
				window.setTimeout( function(){document.getElementById(sid).focus(); }, 0);
			}
		}
	}
	function isNumeric(value) {
    return /^\d+$/.test(value);
	}
	
	function g(){
	document.getElementById("maindiv").style.height = ((document.body.offsetHeight)-180).toString()+"px" ;
	}
</script>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="g()"  >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table  class="txt" cellpadding=3 cellspacing=3>
								<tr>		 
									<TD nowrap align=right >廠別(Xuong)：</TD>
									<TD nowrap>
										<select name=WHSNO   style="width:120px;vertical-align:middle">
											<option value="">----</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if whsno=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											rst.close 
										SET RST=NOTHING  
											%>
										</SELECT>	
										&nbsp;&nbsp;(B)額度(人)&nbsp;&nbsp;
										<input  type="text" name="dfudamt" id="dfudamt" value="<%=dfudamt%>"  onblur="chknum(this.id)"  style="width:100px;vertical-align:middle">
										&nbsp;&nbsp;nguoi/VND
									</TD> 
								</TR>		 
								<tr>
									<TD nowrap   align=right >排序(sap xep)：</TD>									
									<td nowrap>
										<select name="sortby"  onchange="gos()" style="width:160px;vertical-align:middle">
											<option value=""  <%if sortby="" then%>selected<%end if%>>theo sua doi thoi gian異動時間</option>
											<option value="1" <%if sortby="1" then%>selected<%end if%>>theo so the工號</option>
											<option value="2" <%if sortby="2" then%>selected<%end if%>>theo bo phan單位部門</option>				
										</select>
										&nbsp;&nbsp;DV(部門)&nbsp;&nbsp;
										<%SQL="SELECT * FROM BASICCODE WHERE FUNC='groupid' and sys_type<>'AAA' ORDER BY SYS_TYPE "
												SET RST = CONN.EXECUTE(SQL)					
											%>
										<select name="qdv"   style="vertical-align:middle;width:120px" onchange="gos()">
											<option value="" >==select==</option>
											<%WHILE NOT RST.EOF  %>
											<option value="<%=RST("SYS_TYPE")%>" <%if qdv=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>				 
											<%RST.MOVENEXT
												WEND 
											rst.close 
											SET RST=NOTHING  
											conn.close 
											%>
										</select>&nbsp;&nbsp;(Q)搜尋
										<input  type="text" name="query" id="query" value=""  style="width:100px;vertical-align:middle">&nbsp;&nbsp;
										<input type="button" value="(S)k.Tra(查詢)" class="btn btn-sm btn-outline-secondary" style="vertical-align:middle" onclick="gos()">										
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%">		
								<tr height=22  bgcolor="#e4e4e4" class="txt">
									<Td align="center"   nowrap="nowrap">STT</td>
									<Td align="center"   nowrap="nowrap">修改<br>sua</td>
									<Td align="center"  bgcolor="#FFE4E1" nowrap="nowrap">刪除<br>xoa</td>
									<Td align="center" >工號<br>So the</td>
									<Td align="center"  >姓名<br>Ho ten</td>
									<Td align="center" >到職日<br>NVX</td>
									<Td align="center"  >單位<br>Bo Phan</td>			
									<Td align="center"  nowrap="nowrap">(A)人數<br>So nguoi</td>
									<Td align="center"  nowrap="nowrap">(B)額度<br>(人)</td>
									<Td align="center" nowrap="nowrap">(C)免稅<br>總額</td>			
								</tr>
								<%for x = 1 to pagerec 
								if x mod 2 = 0 then 
									wkclr="#FDF5E6"
								else			
									wkclr="#ffffff"
								end if 	 
								if tmprec(currentpage,x,9)="" then types="hidden" else types="checkbox"
								%>
								<Tr bgcolor="<%=wkclr%>">
									<Td align="center"   ><%=x%></td>
									<td  align="center"   >
										<input type="<%=types%>" name="fnop" onclick="fnopchg(<%=x-1%>)">
										<input type="hidden" name="serno" value="<%=tmprec(currentpage,x,9)%>" >
										<input type="hidden" name="op" value="" >
									</td>
									<Td align="center" bgcolor="#FFE4E1"  >
										<input type="<%=types%>" name="ops" onclick="opschg(<%=x-1%>)">										
									</td>
									<Td  ><input type="text" style="width:100%" name="empid" class="inputbox"  onblur="empidchg(<%=x-1%>)" maxlength=6 value="<%=tmprec(currentpage,x,2)%>"></td>
									<Td  ><input type="text" style="width:100%" name="empnam"  value="<%=tmprec(currentpage,x,3)%>" class=readonly8 readonly ></td>
									<Td  ><input type="text" style="width:100%" name="indat"  value="<%=tmprec(currentpage,x,5)%>" class=readonly8 readonly ></td>
									<Td  ><input type="text" style="width:100%" name="gstr"  value="<%=tmprec(currentpage,x,4)%>" class=readonly8 readonly ></td>
									<Td  ><input type="text" name="person_qty" value="<%=tmprec(currentpage,x,6)%>" class="inputbox"  onblur="tot_Mtaxchg1(<%=x-1%>)" style="width:100%;text-align:center"></td>
									<Td  ><input type="text" name="ut_mtax" value="<%=tmprec(currentpage,x,7)%>" class="inputbox"  onblur="tot_Mtaxchg2(<%=x-1%>)" style="width:100%;text-align:right"></td>
									<Td><input type="text" name="tot_Mtax" value="<%=tmprec(currentpage,x,8)%>" class="readonly8" style="width:100%;text-align:right"></td>
								</tr>
								<%next%>								 
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>	    	
									<TD nowrap>		
									<input type="BUTTON" name="send" value="(Y)Confirm" class="btn btn-sm btn-danger" ONCLICK="GO()">
									<input type="BUTTON" name="send" value="(N)Cancel" class="btn btn-sm btn-outline-secondary" onclick="clr()">
									</TD>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			
</form>
<%set conn=nothing	%>
</body>
</html>
 
<!-- #include file="../Include/func.inc" -->

<script language=vbs> 
function f()
	<%=self%>.whsno.focus()
end function 

function clr()
	open "<%=SELF%>.asp" , "_self"
end function 

function opschg(index)	 
	'tot_Mtaxchg(index,1)
	<%=self%>.fnop(index).checked =true 
	<%=self%>.op(index).value="E"
	if  <%=self%>.ops(index).checked then 
		<%=self%>.person_qty(index).value= 0
		<%=self%>.ut_mtax(index).value  = 0
		if trim(<%=self%>.ut_mtax(index).value)<>"" and trim(<%=self%>.person_qty(index).value)<>"" then 
			<%=self%>.tot_Mtax(index).value=formatnumber( cdbl(<%=self%>.person_qty(index).value)*cdbl(<%=self%>.ut_mtax(index).value) , 0)
		else 	
			<%=self%>.tot_Mtax(index).value = 0 
		end if 
	else
		<%=self%>.person_qty(index).value= ""
		<%=self%>.ut_mtax(index).value  = <%=self%>.dfudamt.value
		<%=self%>.person_qty(index).focus()				
	end if 
	
end function 

function  fnopchg(index)
	if <%=self%>.fnop(index).checked then 
		<%=self%>.op(index).value="E" 
	else	
		<%=self%>.op(index).value="" 
	end if	
end function 

function gos()
	<%=self%>.totalpage.value=""
	<%=self%>.action="<%=self%>.fore.asp"
	<%=self%>.submit()
end function 

function tot_Mtaxchg1(index)
	if <%=self%>.person_qty(index).value<>"" then 
		if isnumeric(<%=self%>.person_qty(index).value)=false then 
			alert "請輸入數字,xin danh lai [so] !!"
			<%=self%>.person_qty(index).value=""
			<%=self%>.person_qty(index).focus()
			exit function 
		end if 
	end if 	
	if trim(<%=self%>.ut_mtax(index).value)<>"" and trim(<%=self%>.person_qty(index).value)<>"" then 
		<%=self%>.tot_Mtax(index).value=formatnumber( cdbl(<%=self%>.person_qty(index).value)*cdbl(<%=self%>.ut_mtax(index).value) , 0)
	end if 
end function 

function tot_Mtaxchg2(index)
	if <%=self%>.ut_mtax(index).value<>"" then 
		if isnumeric(<%=self%>.ut_mtax(index).value)=false then 
			alert "請輸入數字,xin danh lai [so] !!"
			<%=self%>.ut_mtax(index).value=""
			<%=self%>.ut_mtax(index).focus()
			exit function 
		else	
			<%=self%>.ut_mtax(index).value=formatnumber(<%=self%>.ut_mtax(index).value,0)
		end if 
	end if  	
	if trim(<%=self%>.ut_mtax(index).value)<>"" and trim(<%=self%>.person_qty(index).value)<>"" then 
		<%=self%>.tot_Mtax(index).value=formatnumber( cdbl(<%=self%>.person_qty(index).value)*cdbl(<%=self%>.ut_mtax(index).value) , 0)
	end if 
end function

function tot_MtaxchgOK(index,a)	
	if a=1 then 
		if <%=self%>.person_qty(index).value<>"" then 
			if isnumeric(<%=self%>.person_qty(index).value)=false then 
				alert "請輸入數字,xin danh lai [so] !!"
				<%=self%>.person_qty(index).value=""
				<%=self%>.person_qty(index).focus()
				exit function 
			end if 
		end if 	
	elseif a=2 then 
		if <%=self%>.ut_mtax(index).value<>"" then 
			if isnumeric(<%=self%>.ut_mtax(index).value)=false then 
				alert "請輸入數字,xin danh lai [so] !!"
				<%=self%>.ut_mtax(index).value=""
				<%=self%>.ut_mtax(index).focus()
				exit function 
			else	
				<%=self%>.ut_mtax(index).value=formatnumber(<%=self%>.ut_mtax(index).value,0)
			end if 
		end if 	
	end if 
	
	if trim(<%=self%>.ut_mtax(index).value)<>"" and trim(<%=self%>.person_qty(index).value)<>"" then 
		<%=self%>.tot_Mtax(index).value=formatnumber( cdbl(<%=self%>.person_qty(index).value)*cdbl(<%=self%>.ut_mtax(index).value) , 0)
	end if 
end function 

function empidchg(index)
	code1=UCase(trim(<%=self%>.empid(index).value))
	if <%=self%>.empid(index).value<>"" then 		
		open "<%=self%>.back.asp?func=chkemp&code1="& code1 &"&index="& index  , "Back"
		'parent.best.cols="50%,50%"
	end if 
end function 
function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go() 
	if <%=self%>.whsno.value="" then 
		alert "請輸入廠別"
		<%=self%>.whsno.focus()
		exit function 
	end if 
	<%=self%>.action="<%=SELF%>.upd.asp"
 	<%=self%>.submit() 
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF 
END FUNCTION
</script> 