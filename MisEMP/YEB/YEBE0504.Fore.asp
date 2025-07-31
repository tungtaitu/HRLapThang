<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
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

q1=trim(request("q1"))
q2=trim(request("q2"))

gTotalPage = 1
PageRec = 10   'number of records per page
TableRec = 15    'number of fields per record    

if q1="" and Q2="" then 
	sql="select ssno, studyName, studygroup, teacher, nw, convert(char(10), d1,111) d1,convert(char(10), d2,111) d2, count(*) as empcnt from empstudy "&_
		"where isnull(status,'')<>'D' and  left(ssno,4)>=year(getdate())-1  group by ssno,  studyName ,  studygroup, teacher , nw,  "&_
		"convert(char(10), d1,111) ,  convert(char(10), d2,111) order by d1 desc, ssno " 
else
	sql="select ssno, studyName, studygroup, teacher, nw, convert(char(10), d1,111) d1,convert(char(10), d2,111) d2, count(*) as empcnt from empstudy "&_
		"where isnull(status,'')<>'D' and  ( left(ssno,4) like '"&q1&"%' or ssno like '"& q1 &"%' ) and   convert(char(6), d1,112) like '"& q2 &"%' and  convert(char(6), d2,112) like '"& q2&"%' "&_
		"group by ssno,  studyName ,  studygroup, teacher , nw,  convert(char(10), d1,111) ,  convert(char(10), d2,111) order by  d1 desc,  ssno" 
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
				tmpRec(i, j, 2) = trim(rs("studyName"))
				tmpRec(i, j, 3) = trim(rs("studygroup"))
				tmpRec(i, j, 4) = trim(rs("teacher"))
				tmpRec(i, j, 5) = trim(rs("nw"))
				tmpRec(i, j, 6) = trim(rs("d1"))
				tmpRec(i, j, 7) = trim(rs("d2"))
				tmpRec(i, j, 8) = trim(rs("empcnt")) 				 		
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
else
	TotalPage = (request("TotalPage"))
	gTotalPage = (request("gTotalPage"))
	'StoreToSession()
	CurrentPage = (request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBE0503")

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
<body onkeydown="enterto()"  onload=f() >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td align=right>編號(年度)</td>
									<td><input type="text" style="width:100px" name=Q1  value="<%=q1%>"></td>
									<td align=right>上課日期(年月)</td>
									<td><input type="text" style="width:100px" name=q2  value="<%=q2%>">(ex:200701)</td>
									<td><input type=button  name=send  class=button value="(S)查詢" onclick=gos() onkeydown=gos()></td>
								</tr>	
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%"> 
								<tr bgcolor=#e4e4e4 height=20>
									<td align=center>編號</td>
									<td align=center>課程名稱</td>
									<td align=center>訓練單位</td>
									<td align=center>講師</td>
									<td align=center>類別</td>
									<td align=center>上課日期(起)</td>
									<td align=center>上課日期(迄)</td>
									<td align=center>受訓<BR>人數</td>
						 
								</tr> 
								<%for CurrentRow = 1 to pagerec        
									 if CurrentRow mod 2 = 0 then
										wkcolor="LightGoldenrodYellow"
									 else
										wkcolor="#ffffff"
									 end if
								%>        	
									<tr bgcolor="<%=wkcolor%>" height=18>	
										<td class=txt align=center>
											<a href="<%=self%>.foregnd.asp?ssno=<%=tmpRec(CurrentPage, CurrentRow, 1)%>&d1=<%=tmpRec(CurrentPage, CurrentRow, 6)%>&D2=<%=tmpRec(CurrentPage, CurrentRow, 7)%>">
												<font color=blue><%=tmpRec(CurrentPage, CurrentRow, 1)%></font>
											</a>
										</td>
										<td>
											<a href="<%=self%>.foregnd.asp?ssno=<%=tmpRec(CurrentPage, CurrentRow, 1)%>&d1=<%=tmpRec(CurrentPage, CurrentRow, 6)%>&D2=<%=tmpRec(CurrentPage, CurrentRow, 7)%>">
											<font color=blue><%=tmpRec(CurrentPage, CurrentRow, 2)%></font>
											</a>
										</td>
										<td><%=tmpRec(CurrentPage, CurrentRow, 3)%></td>
										<td ><%=tmpRec(CurrentPage, CurrentRow, 4)%></td>
										<td align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%></td>
										<td align=center><%=tmpRec(CurrentPage, CurrentRow, 6)%></td>
										<td align=center><%=tmpRec(CurrentPage, CurrentRow, 7)%></td>
										<td align=center><%=tmpRec(CurrentPage, CurrentRow, 8)%></td>
										<input type=hidden name=ssno value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
										<input type=hidden name=D1 value="<%=tmpRec(CurrentPage, CurrentRow, 6)%>">
										<input type=hidden name=D2 value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>"> 
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
	
<table width=500  ><tr><td >
	
    
	     

</td></tr></table> 

</body>
</html>


<script language=vbs> 
function deldata(index)
	'alert index 
	'alert <%=self%>.ssno(index).value
	'alert <%=self%>.d1(index).value
	'alert <%=self%>.d2(index).value
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