<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()   
self="YEBE0504"   

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

gTotalPage = 1
PageRec = 20    'number of records per page
TableRec = 15    'number of fields per record     

ssno = request("ssno")
d1 = request("d1")
d2 = request("d2")
sql="select b.job, b.jstr, b.gstr, b.empnam_cn, b.empnam_vn , c.amt, c.dm as plandm, convert(char(10),d1,111) as DD1, convert(char(10),d2,111) as DD2,  a.* from "&_
	"( select * from empstudy where isnull(status,'')<>'D' and ssno='"& ssno &"' and convert(char(10),d1,111)='"& d1 &"' and convert(char(10),d2,111)='"& D2 &"' ) a "&_
	"join (select empid , empnam_cn , empnam_vn, jstr, job , gstr from view_empfile ) b on b.empid = a.empid "&_
	"join (select *   from studyplan ) c on c.ssno = a.ssno  " 
'response.write sql 	
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1 	 	  		
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQL, conn, 3, 3  
	IF NOT RS.EOF THEN
		PageRec = rs.RecordCount 
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then	
				F_ssno=trim(rs("ssno"))
				F_studyName=trim(rs("studyName"))
				F_studygroup=trim(rs("studygroup"))
				F_teacher=trim(rs("teacher"))
				F_nw=trim(rs("nw"))
				F_D1=trim(rs("dd1"))
				F_D2=trim(rs("dd2"))
				F_tim1=trim(rs("tim1"))
				F_tim2=trim(rs("tim2"))
				F_whour=trim(rs("whour"))
				F_memo=trim(rs("memo"))
				F_amt=trim(rs("amt"))
				F_dm=trim(rs("plandm"))
				
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("groupid"))
				tmpRec(i, j, 3) = trim(rs("empnam_cn"))
				tmpRec(i, j, 4) = trim(rs("empnam_vn"))
				tmpRec(i, j, 5) = trim(rs("pzjno"))				
				tmpRec(i, j, 6) = trim(rs("whsno"))
				tmpRec(i, j, 7) = trim(rs("country"))
				tmpRec(i, j, 8) = trim(rs("aid"))
				tmpRec(i, j, 9) = trim(rs("jstr"))
				tmpRec(i, j, 10) = trim(rs("pjsts"))
				tmpRec(i, j, 11) = trim(rs("fensu"))
				tmpRec(i, j, 12) = trim(rs("samt"))
				tmpRec(i, j, 13) = trim(rs("dm"))
				tmpRec(i, j, 14) = trim(rs("gstr"))
				rs.movenext
			else
				exit for			
			end if 
	
			if rs.EOF then
				rs.Close
				Set rs = nothing
				exit for
			 end if
		next
	next 
	Session("YEBE0503") = tmpRec 
end if 	
'response.write pagerec

conn.close
set conn=nothing 
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
    <%=self%>.fensu(0).focus()   
    '<%=self%>.country.SELECT()
end function  

function gob()
	open "<%=self%>.asp" , "_self"
end function    
-->
</SCRIPT>   
</head> 
<body   onkeydown="enterto()"  onload=f() >
<form name="<%=self%>" method="post"  >
<input type=hidden name=pagerec value="<%=pagerec%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr bgcolor=#e4e4e4 >
									<td align=right width=100>上課日期<br>Ngay huan luyen</td>
									<td width=220 ><b><%=F_d1%> ~ <%=F_d2%></b>
										<input type=hidden id=1 name="D1" size=11 class=readonly readonly onblur=date_change(1) value="<%=F_d1%>" title="上課日期(起)"  >  
										<input type=hidden id=2 name="D2" size=11 class=readonly readonly onblur=date_change(2) value="<%=F_d2%>" title="上課日期(迄)"  >
									</td>
									<td align=right width=90 >類別<br>Loai huan luyen</td> 
									<td  >
										<select name=nw class=inputbox disabled >            		
											<option value="N" <%if f_nw="N" then%>selected<%end if%>>內(Nội) </option>
											<option value="W" <%if f_nw="W" then%>selected<%end if%>>外(Ngoại)</option>
										</select>
									</td>            
								</tr>
								<!--tr  bgcolor=#e4e4e4 >  
									<td align=right>上課時間<br>Giao huan luyen</td>
									<td  colspan=3>
										Tu <input id=3 name="tim1" size=6 class=inputbox onblur=timchg(1) maxlength=5 value="<%=F_tim1%>" title="上課時間(起)"  > ~ 
										Den <input id=4 name="tim2" size=6 class=inputbox onblur=timchg(2) maxlength=5 value="<%=F_tim2%>" title="上課時間(迄)"  > 共
										<input name="whour" size=4 class=inputbox value="<%=F_whour%>" onblur=whourchg()> (Gio/Hour)
									</td>
								</tr-->
								<tr  bgcolor=#e4e4e4 >
									<td align=right >課程名稱<br>Ten Giao Trinh </td>
									<td  colspan=3><%=f_ssno%>&nbsp;<%=F_studyname%>
										<input type=hidden name="ssno" class=readonly value="<%=f_ssno%>" size=10 readonly  ><br>
										<input type=hidden  id=5 name="studyName" size=51 value="<%=F_studyname%>" class=readonly readonly   title="課程名稱(Ten Giao Trinh)">                
									</td>
								</tr>
								<tr  bgcolor=#e4e4e4 >
									<td align=right  >訓練單位<br>Don vi huan luyen</td>
									<td   ><%=f_studyGroup%><input type=hidden name="studyGroup" value="<%=f_studyGroup%>" size=51 class=inputbox></td>
									<td align=right  >講師<br>Giang vien</td>
									<td   ><%=f_teacher%> <input type=hidden  name="teacher" value="<%=f_teacher%>" size=51 class=inputbox></td>
								</tr>  
						 
								<tr  bgcolor=#e4e4e4 >
									<td align=right  >費用<br></td>
									<td  >
										<input name=amt size=11  value="<%=f_amt%>" class=readonly readonly  > 
										<select name=dm class=txt8 disabled > 
											<option value=""></option>
											<option value="VND" <%if f_dm="VND" then%>selected<%end if%>>VND</option>
											<option value="USD" <%if f_dm="USD" then%>selected<%end if%>>USD</option>
										</select>
									</td>         
									<td align=right  >備註<br>Ghi Chu</td>
									<td   >  <%=f_memo%> 
									</td>
								</tr> 
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table table-bordered table-sm bg-white text-secondary"> 
								<tr bgcolor=#e4e4e4 height=20>        	
									<td align=center>工號</td>
									<td align=center >單位</td>
									<td align=center>姓名</td>
									<td align=center>職稱</td>
									<td align=center>評分</td>
									<td align=center>評鑑結果</td>
									<td align=center>費用</td>
									<td align=center>幣別</td>
									<td align=center>證照號碼</td>   
									
								</tr>
								<%for t = 1 to PageRec%>
									<tr>	 
										<td><input name=empid size=5 class=inputbox value="<%=tmpRec(CurrentPage, t, 1)%>" <%if tmpRec(CurrentPage, t, 1)<>"" then%>readonly style='background-color:lightYellow'<%else%> onblur="empidchg(<%=t-1%>)" <%end if %>></td>
										<td>
											<input name=gstr size=6 class=readonly  readonly style='background-color:lightYellow' value="<%=tmpRec(CurrentPage, t, 14)%>" >
											<input type=hidden name=groupid   value="<%=tmpRec(CurrentPage, t, 2)%>" >
										</td>
										<td><input name=empname size=15 class=readonly readonly style='background-color:lightYellow'  value="<%=tmpRec(CurrentPage, t, 3)&tmpRec(CurrentPage, t, 4)%>"  ></td>
										<td><input name=jobstr size=10 class=readonly readonly  value="<%=tmpRec(CurrentPage, t, 9)%>" ></td>
										<td><input name=fensu size=4 class=inputbox8    value="<%=tmpRec(CurrentPage, t, 11)%>" ></td>
										<td>
											<select name=pjsts class=txt >
												<option value=""></option>
												<option value="Y" <%if tmpRec(CurrentPage, t, 10)="Y" then %>selected<%end if%>>OK</option>
												<option value="N" <%if tmpRec(CurrentPage, t, 10)="N" then %>selected<%end if%>>No OK</option>
											</select>        			
										</td>
										<td><input name=samt size=10 class=inputboxr    value="<%=tmpRec(CurrentPage, t, 12)%>" ></td>
										<td>
											<select name=pdm class=txt    > 
												<option value=""></option>
												<option value="VND" <%if tmpRec(CurrentPage, t, 13)="VND" then%>selected<%end if%>>VND</option>
												<option value="USD" <%if tmpRec(CurrentPage, t, 13)="USD" then%>selected<%end if%>>USD</option>
											</select>        			
										</td>
										<td>
											<input name=pzjno size=15 class=readonly readonly value="<%=tmpRec(CurrentPage, t, 5)%>">
											<input type=hidden name=whsno size=15 class=inputbox8  value="<%=tmpRec(CurrentPage, t, 6)%>" >
											<input type=hidden name=country size=15 class=inputbox8  value="<%=tmpRec(CurrentPage, t, 7)%>" >
											<input type=hidden name=aid size=1 class=inputbox8  value="<%=tmpRec(CurrentPage, t, 8)%>" >        		
										</td>        		
									 
									</tr> 
								<%next%>
								<tr>
									<td colspan=9 align=center>count:<%=recordInDB%></td>
								</tr>
									<input name=empid  value="" type=hidden>
									<input name=gstr  value="" type=hidden>
									<input name=groupid  value="" type=hidden>
									<input name=empname  value="" type=hidden>
									<input name=jobstr  value="" type=hidden>
									<input name=fensu  value="" type=hidden>
									<input name=pjsts  value="" type=hidden>
									<input name=samt  value="" type=hidden>
									<input name=pdm  value="" type=hidden>
									<input name=pzjno  value="" type=hidden>
									<input name=whsno  value="" type=hidden>
									<input name=country  value="" type=hidden>
									<input name=aid  value="" type=hidden>         
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr >
									<td align=center>
										<input type=button  name=btm class=button value="確   認" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class=button value="取   消">
										<input type=reset  name=btm class=button value="回主畫面(Main)" onclick='gob()'>
									</td>
								</tr>   
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	
<table width=500  ><tr><td >
        
    <hr size=0  style='border: 1px dotted #999999;' align=left width=600>       
    
        

</td></tr></table> 

</body>
</html>
<script language=vbs>

function delchg(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="del"
	else
		<%=self%>.op(index).value=""
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
				'Elm.value=ucase(Elm.value)
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
    parent.best.cols="50%,50%" 
    
    'open "Getempdata.asp?pself="& "<%=self%>" &"&index=" & index &"&ncols="& ncols , "Back"   
end function 
</script> 