<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->

<%
Set conn = GetSQLServerConnection()   
self="YEBE0505"  


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

q1=request("q1")
q2=request("q2")
q3=request("q3")
q4=request("q4")
q5=request("q5")
fg=request("fg")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 25    'number of fields per record    

if q1="" and Q2="" and q3="" and q4="" and Q5=""  and fg="" then 	
	sql="select c.cstr, c.wstr, c.gstr, c.empnam_cn, c.empnam_vn, b.studyname planName, "&_
		"convert(char(6), d1,112) as studyYM,convert(char(10), d1,111) dd1, convert(char(10),d2,111) dd2 ,  a.*  from  "&_ 
		"(select *   from  empstudy  where  isnull(status,'')<>'D' and ssno='*'  )  a "&_
		"join (select  * from studyplan ) b on b.ssno = a.ssno  "&_
		"join (select * from view_empfile ) c on  c.empid = a.empid order by a.ssno, a.d1 , a.empid  "
 
else
	sql="select c.cstr, c.wstr, c.gstr, c.empnam_cn, c.empnam_vn, b.studyname planName, "&_
		"convert(char(6), d1,112) as studyYM, convert(char(10), d1,111) dd1, convert(char(10),d2,111) dd2 , a.*  from  "&_ 
		"(select *   from  empstudy  where  isnull(status,'')<>'D'  )  a "&_
		"join (select  * from studyplan ) b on b.ssno = a.ssno  "&_
		"join (select * from view_empfile ) c on  c.empid = a.empid  "&_
		"where a.empid like '"& Q1 &"%' and a.groupid like '"& Q2 &"%' and a.country like '"& Q3 &"%' and a.ssno like '"& Q4 &"%'  "&_
		"and convert(char(6), d1,112) like '"& Q5 &"%'  order by a.ssno, a.d1 , a.empid "
end if 		

'response.write sql 
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1 	 	  		
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQL, conn, 3, 3  
	IF NOT RS.EOF THEN
		'PageRec = rs.RecordCount 
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then	
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("ssno"))
				tmpRec(i, j, 2) = trim(rs("planName"))
				tmpRec(i, j, 3) = trim(rs("studygroup"))
				tmpRec(i, j, 4) = trim(rs("teacher"))
				tmpRec(i, j, 5) = trim(rs("nw"))
				tmpRec(i, j, 6) = trim(rs("dd1"))
				tmpRec(i, j, 7) = trim(rs("dd2"))
				tmpRec(i, j, 8) = trim(rs("empid"))
				tmpRec(i, j, 9) = trim(rs("empnam_cn"))
				tmpRec(i, j, 10) = trim(rs("empnam_vn"))
				tmpRec(i, j, 11) = trim(rs("groupid"))
				tmpRec(i, j, 12) = trim(rs("country"))
				tmpRec(i, j, 13) = trim(rs("gstr"))
				tmpRec(i, j, 14) = trim(rs("cstr"))
				tmpRec(i, j, 15) = trim(rs("tim1"))&"~"&trim(rs("tim2"))
				tmpRec(i, j, 16) = trim(rs("tim2"))
				tmpRec(i, j, 17) = trim(rs("whour")) 
				tmpRec(i, j, 18) = trim(rs("pjsts")) 
				tmpRec(i, j, 19) = trim(rs("pzjno")) 
				if rs("dd1")=rs("dd2") then 
					tmpRec(i, j, 20) = trim(rs("dd1"))&"~"&trim(rs("dd2"))
				else
					tmpRec(i, j, 20) = trim(rs("dd1"))&"~"&trim(rs("dd2"))
				end if 	
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
	Session("YEBE0505") = tmpRec
else
	TotalPage = (request("TotalPage"))
	gTotalPage = (request("gTotalPage"))
	'StoreToSession()
	CurrentPage = (request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBE0505")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(gTotalPage) then
			     CurrentPage = CurrentPage + 1
			  else
			  	 CurrentPage = TotalPage
		      end if
	     Case "END"
		      CurrentPage = gTotalPage
	     Case Else
		      CurrentPage = 1
	end Select 	
end if 	  
	
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<!-- #include file="../Include/func.inc" -->
<SCRIPT   LANGUAGE=vbscript>
 
function f()
    <%=self%>.q1.focus()   
    '<%=self%>.country.SELECT()
end function   

function gos()
	<%=self%>.totalpage.value="0"
	<%=self%>.action="<%=self%>.Fore.asp"
	<%=self%>.submit()
end function   
 
</SCRIPT>   
</head> 
<body  onkeydown="enterto()"  onload=f() >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=fg value="1">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<TD width=80 align=right>國籍:</TD>
									<TD width=80>
										<select name=Q3 class=txt8 style='width:75px'   >
											<option value="">全部 </option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if Q3=RST("SYS_TYPE") then%>selected<%END IF%> ><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											rst.close
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>							
									<TD  width=80 align=right >組/部門:</TD>
									<TD width=100 >
										<select name=Q2  class=txt8    >
										<option value="" selected >全部 </option>
										<%
										SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
										'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
										SET RST = CONN.EXECUTE(SQL)
										'RESPONSE.WRITE SQL 
										WHILE NOT RST.EOF  
										%>
										<option value="<%=RST("SYS_TYPE")%>" <%if Q2=RST("SYS_TYPE") then%>selected<%END IF%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
										<%
										RST.MOVENEXT
										WEND 
										rst.close
										%>
										</SELECT>
										<%SET RST=NOTHING %>
									</td>	
									<td align=right>工號:</td>
									<td><input type="text" style="width:100px" name=Q1  value="<%=q1%>" size=6></td>
								</tr>
								<tr>		
									<td align=right>課程編號:</td>
									<td><input type="text" style="width:100px" name=Q4  value="<%=q4%>" size=10></td>
									<td align=right>上課年月:</td>
									<td><input type="text" style="width:100px" name=Q5  value="<%=q5%>" size=10></td>		
									<td align=center colspan=2>
										<input type=button name=send value="(S)查詢" onclick="gos()" onkeydown="gos()" class=button>
									</td>
								</tr>	
							</table>
							<%conn.close
							set conn=nothing
							%>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%"> 
								<tr bgcolor=#e4e4e4 height=20>
									<td align=center width=50 NOWRAP>單位</td>
									<td align=center width=50 NOWRAP>工號</td>
									<td align=center width=120 NOWRAP>員工姓名</td>    		
									<td align=center width=150 NOWRAP>課程名稱</td>
									<td align=center width=125 NOWRAP>上課日期</td>
									<td align=center width=75 NOWRAP>時間</td>
									<td align=center width=40 NOWRAP>時數</td>
									<td align=center width=80 NOWRAP>證照號碼</td>
									<td align=center width=40 NOWRAP>評鑑<BR>結果</td>
								</tr> 
								<%for CurrentRow = 1 to pagerec        
									 if CurrentRow mod 2 = 0 then
										wkcolor="LightGoldenrodYellow"
									 else
										wkcolor="#ffffff"
									 end if
								%>        	
									<tr bgcolor="<%=wkcolor%>" height=20 class=txt8>	
										<td  align=center>        			
											<%=tmpRec(CurrentPage, CurrentRow, 11)%><BR>
											<%=tmpRec(CurrentPage, CurrentRow, 13)%>
										</td>
										<td align=center>
											<%=tmpRec(CurrentPage, CurrentRow, 8)%>
										</td>
										<td><%=tmpRec(CurrentPage, CurrentRow, 9)%><BR><%=tmpRec(CurrentPage, CurrentRow, 10)%></td>        		
										<td  ><%=tmpRec(CurrentPage, CurrentRow, 2)%></td>
										<td  align=center ><%=tmpRec(CurrentPage, CurrentRow, 20)%></td>
										<td  align=center ><%=tmpRec(CurrentPage, CurrentRow, 15)%></td>
										<td align=center><%=tmpRec(CurrentPage, CurrentRow, 17)%></td>        		
										<td align=center><%=tmpRec(CurrentPage, CurrentRow, 18)%></td>
										<td align=center><%=tmpRec(CurrentPage, CurrentRow, 19)%></td>        		
									</tr>
								<%next%>	    	
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr>
									<td align="CENTER" height=40 WIDTH=60%>
									<% If CurrentPage > 1 Then %>
										<input type="submit" name="send" value="FIRST" class=button>
										<input type="submit" name="send" value="BACK" class=button>
									<% Else %>
										<input type="submit" name="send" value="FIRST" disabled class=button>
										<input type="submit" name="send" value="BACK" disabled class=button>
									<% End If %>
									<% If cint(CurrentPage) < cint(gTotalPage) Then %>
										<input type="submit" name="send" value="NEXT" class=button>
										<input type="submit" name="send" value="END" class=button>
									<% Else %>
										<input type="submit" name="send" value="NEXT" disabled class=button>
										<input type="submit" name="send" value="END" disabled class=button>
									<% End If %>
									<FONT CLASS=TXT8>&nbsp;&nbsp;PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%></FONT>
									</TD> 
								</TR> 
								</TABLE>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

</body>
</html>


<script language=vbs> 
function deldata(index)
	f_ssno=trim(<%=self%>.ssno(index).value)
	f_d1=trim(<%=self%>.d1(index).value)
	f_d2=trim(<%=self%>.d2(index).value)
	if confirm("確定要刪除此筆資料?"&chr(13)&"Xoa Tu Lieu??",64) then 
		open "<%=self%>.upd.asp?flag=del&ssno="&f_ssno&"&d1="&f_d1&"&d2=" & f_d2 , "Back" 
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
				Elm.value=escape(ucase(Elm.value))
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
		parent.best.cols="70%,30%"
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