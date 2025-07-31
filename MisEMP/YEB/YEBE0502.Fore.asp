<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->

<%
'Set conn = GetSQLServerConnection()   
self="YEBE0502"  


nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
    calcmonth = year(date()-1)&"12"     
else
    calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if     

if day(date())<=11 then 
    if month(date())="01" then  
        calcmonth = year(date()-1)&"12" 
    else
        calcmonth =  year(date())&right("00"&month(date())-1,2)   
    end if      
else
    calcmonth = nowmonth 
end if 
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
    <%=self%>.D1.focus()   
    '<%=self%>.country.SELECT()
end function    
-->
</SCRIPT>   
</head> 
<body  onkeydown="enterto()"  onload=f() >
<form name="<%=self%>" method="post"  >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">
							<table id="myTableForm" width="70%">
								<tr>
									<td align=right nowrap style="width:10%">上課日期<br><font class="txt8">Ngay huan luyen</font></td>
									<td colspan=3>
										<input id=1 name="D1" size=11  onblur=date_change(1) title="上課日期(起)">~
										<input id=2 name="D2" size=11  onblur=date_change(2) title="上課日期(迄)">
									</td>									
								</tr>
								<tr>  
									<td align=right nowrap>上課時間<br><font class="txt8">Giao huan luyen</font></td>
									<td colspan=3>
										<input id=3 name="tim1" size=6  onblur=timchg(1) maxlength=5 title="上課時間(起)">~
										<input id=4 name="tim2" size=6  onblur=timchg(2) maxlength=5 title="上課時間(迄)">共
										<input name="whour" size=4  onblur=whourchg()>(Gio/Hour)
									</td>
								</tr>
								<tr>
									<td align=right nowrap><a href="vbscript:gotstudyplan()"><font color=blue>課程名稱<br><font class="txt8">Ten Giao Trinh</font></font></a></td>
									<td colspan=3>
										<input name="ssno"  size=10 style="width:100px">
										<input id=5 name="studyName" size=51  ondblclick="gotstudyplan()" title="課程名稱(Ten Giao Trinh)">
									</td>
								</tr>
								<tr>
									<td align=right  nowrap style="width:10%">訓練單位<br><font class="txt8">Don vi huan luyen</font></td>
									<td style="width:40%"><input type="text" name="studyGroup" ></td>								
									<td align=right nowrap style="width:10%">講師<br><font class="txt8">Giang vien</font></td>
									<td style="width:40%"><input type="text" name="teacher"></td>
								</tr> 
								<tr>
									<td align=right nowrap>類別<br><font class="txt8">Loai huan luyen</font></td>
									<td >
										<select name=nw >            		
											<option value="N">內(Nội) </option>
											<option value="W">外(Ngoại)</option>
										</select>
									</td>
									<td align=right>證照</td>
									<td>
										<select name=pzj >
											<option value="">無(Ko.)</option>
											<option value="Y">有(Co.)</option>
										</select>
									</td>
								</tr> 
								<tr>
									<td align=right>費用<br></td>
									<td colspan=3>
										<input name=amt size=11>
										<select name=dm style="width:100px">
											<option value=""></option>
											<option value="VND">VND</option>
											<option value="USD">USD</option>
										</select>												
									</td>									
								</tr> 
								<tr>
									<td align=right>備註<br><font class="txt8">Ghi Chu</font></td>
									<td colspan=3>  
										<input type="text" style="width:98%" name="memo">
									</td>
								</tr> 
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td  align="center">
							<table id="myTableGrid" width="98%"> 
								<tr class="bg-gray text-black" height="35px">
									<td align=center>工號</td>
									<td align=center >單位</td>
									<td align=center>姓名</td>
									<td align=center>證照號碼</td>
									<td width=5 bgcolor=#ffffff></td>
									<td align=center >工號</td>
									<td align=center>單位</td>
									<td align=center>姓名</td>
									<td align=center>證照號碼</td>
								</tr>
								<%for t = 1 to 20%>
									<%if t mod 2 = 1 then %><tr><%end if%>
										<td><input type="text" style="width:98%" name=empid onblur="empidchg(<%=t-1%>)"></td>
										<td><input type="text" style="width:98%;background-color:lightYellow" name=groupid  readonly ></td>
										<td><input type="text" style="width:98%;background-color:lightYellow" name=empname  readonly></td>
										<td>
											<input type="text" style="width:98%" name=pzjno  >
											<input type=hidden name=whsno  >
											<input type=hidden name=country  >									
										</td>
										<%if t mod 2 = 1 then %><td width=5></td><%end if%>
									<%if t mod 2 = 0 then %></tr><%end if%>
								<%next%>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt">
								<tr>
									<td align=center>
										<input type=button  name=btm class="btn btn-sm btn-danger" value="確 認" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取 消">
									</td>
								</tr>   
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

</body>
</html>


<script language=vbs>
function fcchg(index)
	if <%=self%>.fc(0).checked=true then 
		<%=self%>.fc(1).checked=false 
		<%=self%>.nw.value="N"
	end if 
	if <%=self%>.fc(1).checked=true then 
		<%=self%>.fc(0).checked=false 
		<%=self%>.nw.value="W"
	end if 
end function 
function dataclick(a)
    if a = 1 then       
        open "empbasic/empbasic.asp" , "_self"
    elseif a = 2 then       
        open "empfile/empfile.asp" , "_self"
    elseif a = 3 then       
        open "empworkHour/empwork.asp" , "_self"    
    elseif a = 4 then       
        open "holiday/empholiday.asp" , "_self" 
    elseif a = 5 then       
        open "AcceptCaTime/main.asp" , "_self"              
    elseif a = 6 then       
        open "../report/main.asp" , "_self"     
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
	dim Elm
	for each Elm in <%=self%>
		'alert  Elm.name
		select case Elm.name
			case "D1","D2","tim1","tim2","whour","studyName" 
				if trim(Elm.value)="" then 
					alert Elm.title & " 必須輸入資料(hay nhap vao tu lieu)"
					Elm.focus()
					exit function 					
				end if 	
			case else
				Elm.value=replace(Elm.value,"'","′")
				Elm.value=replace(Elm.value,"""","”")
				Elm.value=ucase(Elm.value)
		end select
	next 
	
    <%=self%>.action="<%=self%>.upd.asp"
    <%=self%>.submit() 
end function   
    
function timchg(a)
	if a = 1 then 
		if <%=self%>.tim1.value<>"" then 
			<%=self%>.tim1.value=left(<%=self%>.tim1.value,2)&":"&right(<%=self%>.tim1.value,2)
		end if 
	elseif a = 2 then 
		if <%=self%>.tim2.value<>"" then 
			<%=self%>.tim2.value=left(<%=self%>.tim2.value,2)&":"&right(<%=self%>.tim2.value,2)
		end if 
	end if  				
end function 

function whourchg()
	if <%=self%>.whour.value<>"" then 
		if isnumeric(<%=self%>.whour.value)=false then 
			alert "請輸入數值!!" 
			<%=self%>.whour.value=""
			<%=self%>.whour.focus()
			exit function 
		end if
	end if 	 
			
end function

function empidchg(index)
	empidstr = trim(<%=self%>.empid(index).value )
	if empidstr<>"" then 
		open "<%=self%>.back.asp?func=chkemp&index=" & index &"&code1=" & empidstr , "Back" 
		'parent.best.cols="70%,30%"
	end if 
end function 


'*******檢查日期*********************************************
FUNCTION date_change(a) 

if a=1 then
    INcardat = Trim(<%=self%>.D1.value)         
elseif a=2 then
    INcardat = Trim(<%=self%>.D2.value)
end if      
            
IF INcardat<>"" THEN
    ANS=validDate(INcardat)
    IF ANS <> "" THEN
        if a=1 then
            Document.<%=self%>.D1.value=ANS         
            if Document.<%=self%>.D2.value="" then 
            	Document.<%=self%>.D2.value=ans
            end if 	
        elseif a=2 then
            Document.<%=self%>.D2.value=ANS                 
        end if      
    ELSE
        ALERT "EZ0067:輸入日期不合法 !!" 
        if a=1 then
            Document.<%=self%>.D1.value=""
            Document.<%=self%>.D1.focus()
        elseif a=2 then
            Document.<%=self%>.D2.value=""
            Document.<%=self%>.D2.focus()
        end if      
        EXIT FUNCTION
    END IF
         
ELSE
    'alert "EZ0015:日期該欄位必須輸入資料 !!"      
    EXIT FUNCTION
END IF 
   
END FUNCTION 

function gotstudyplan()
    ncols="studyGroup" 
    open "getstudyPlan.asp?pself="& "<%=self%>" &"&ncols="& ncols , "Back" 
    parent.best.cols="70%,30%" 
    
    'open "Getempdata.asp?pself="& "<%=self%>" &"&index=" & index &"&ncols="& ncols , "Back"   
end function 
</script> 