<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<%
'Set conn = GetSQLServerConnection()   
self="YEBE05"  


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
    <%=self%>.workdat.focus()   
    '<%=self%>.country.SELECT()
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"  onload=f() >
<form name="<%=self%>" method="post"  >
<table width="460" border="0" cellspacing="0" cellpadding="0">
    <tr><TD>
    <img border="0" src="../image/icon.gif" align="absmiddle">
    <%=session("pgname")%></TD></tr>
</table>
<hr size=0  style='border: 1px dotted #999999;' align=left width=500>
<table width="460" border="0" cellspacing="0" cellpadding="0" class=txt99>
    <tr>
        <TD width=160><a href="<%=self%>.Plan.asp">教育訓練計劃(Ke hoach)</a></TD>
        <TD width=80><a href="<%=self%>.Fore.asp">查詢(K.Tra)</a></TD>
        <TD width=80><a href="<%=self%>.New.asp">新增(Moi)</a></TD>
        <TD width=120><a href="<%=self%>.edit.asp">資料刪修(Xoa/Doi)</a></TD>
    </tr>
</table>
<hr size=0  style='border: 1px dotted #999999;' align=left width=500>       
<table width=500  ><tr><td >
    <table width=450 class=txt99>
        <tr bgcolor=#e4e4e4 >
            <td align=right width=120>上課日期<br>Ngay huan luyen</td>
            <td><input name="workdat" size=11 class=inputbox></td>
        </tr>
        <tr  bgcolor=#e4e4e4 >  
            <td align=right>上課時間<br>Giao huan luyen</td>
            <td>
                Tu <input name="tim1" size=6 class=inputbox> ~ 
                Den <input name="tim2" size=6 class=inputbox> 共
                <input name="whour" size=4 class=inputbox> (Gio/Hour)
            </td>
        </tr>
        <tr  bgcolor=#e4e4e4 >
            <td align=right  >課程名稱<br>Ten Giao Trinh</td>
            <td>
                <input name="studyName" size=51 class=inputbox ondblclick="gotstudyplan()">
                <input type=hidden name="ssno" >
            </td>
        </tr>
        <tr  bgcolor=#e4e4e4 >
            <td align=right  >訓練單位<br>Don vi huan luyen</td>
            <td><input name="studyGroup" size=51 class=inputbox></td>
        </tr> 
        <tr  bgcolor=#e4e4e4 >
            <td align=right  >講師<br>Giang vien</td>
            <td><input name="teacher" size=51 class=inputbox></td>
        </tr> 
        <tr  bgcolor=#e4e4e4 >
            <td align=right  >類別<br>Loai huan luyen</td>
            <td><input type=checkbox name="fc">內(Nội)  
                <input type=checkbox name="fc">外 (Ngoại)
                <input type=hidden value="" name="nw">
            </td>
        </tr> 
            <tr  bgcolor=#e4e4e4 >
            <td align=right  >費用<br></td>
            <td>
                <input name=amt size=11 class=inputbox > 
                <select name=dm class=txt8>
                    <option value=""></option>
                    <option value="VND">VND</option>
                    <option value="USD">USD</option>
                </select>
            </td>
        </tr> 
    </table>    
    <hr size=0  style='border: 1px dotted #999999;' align=left width=500>       
    <table width=500 class=txt99 > 
        <tr bgcolor=#e4e4e4>
            <td align=center>工號</td>
            <td align=center >單位</td>
            <td align=center>姓名</td>
            <td width=5 bgcolor=#ffffff></td>
            <td align=center >工號</td>
            <td align=center>單位</td>
            <td align=center>姓名</td>
        </tr>
        <%for t = 1 to 30%>
        	<%if t mod 2 = 1 then %><tr><%end if%>
        		<td><input name=empid size=6 class=inputbox ></td>
        		<td><input name=groupid size=6 class=inputbox ></td>
        		<td><input name=empname size=20 class=inputbox8 ></td>
        		<%if t mod 2 = 1 then %><td width=5></td><%end if%>
        	<%if t mod 2 = 0 then %></tr><%end if%>
        <%next%>
    </table>
    <table width=450 align=center>
        <tr >
            <td align=center>
                <input type=button  name=btm class=button value="確   認" onclick="go()" onkeydown="go()">
                <input type=reset  name=btm class=button value="取   消">
            </td>
        </tr>   
    </table>    

</td></tr></table> 

</body>
</html>


<script language=vbs>
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
    <%=self%>.action="<%=self%>.ForeGnd.asp"
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

function gotstudyplan()
    ncols="studyGroup" 
    open "getstudyPlan.asp?pself="& "<%=self%>" &"&ncols="& ncols , "Back" 
    parent.best.cols="50%,50%" 
    
    'open "Getempdata.asp?pself="& "<%=self%>" &"&index=" & index &"&ncols="& ncols , "Back"   
end function 
</script> 