<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

SELF = "YEBE0301B"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

firstday = year(date())&"/"&right("00"&month(date()),2)&"/01" 
nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 
  
empautoid = TRIM(REQUEST("empautoid"))  
bdy=""
SQL="SELECT * FROM view_empfile where ISNULL(STATUS,'')<>'D' AND  autoid='"& empautoid &"' "

'RESPONSE.WRITE SQL
'RESPONSE.END
RS.OPEN SQL , CONN, 3, 3
IF NOT RS.EOF THEN
	empautoid = TRIM(RS("AUTOID"))
	emptype = TRIM(RS("emptype"))
	EMPID=TRIM(RS("EMPID"))	'員工編號
	INDAT=TRIM(RS("nindat"))	'到職日
	WHSNO=TRIM(RS("WHSNO"))	'廠別
	UNITNO=TRIM(RS("UNITNO"))	'處/所
	GROUPID=TRIM(RS("GROUPID"))	'組/部門
	ZUNO=TRIM(RS("ZUNO"))	'單位

	EMPNAM_CN=TRIM(RS("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RS("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RS("COUNTRY"))	'國籍
	COUNTRYDESC=TRIM(RS("cstr"))	'國籍
	BYY=(TRIM(RS("BYY"))) '年(生日)
	BMM=(RS("BMM"))	'月(生日)
	BDD=(RS("BDD"))	'日(生日)
	if byy<>"" then 
		bdy=BYY 
	end if	
	if bmm<>"" then 
		bdy=bdy & "/"&right(bmm,2)
	end if	
	if bdd<>"" then 		
		bdy=bdy & "/"&right(bdd,2)	
	end if 	
	AGES=TRIM(RS("AGES"))	'年齡
	SEX=TRIM(RS("SEX"))	'性別
	JOB=TRIM(RS("JOB"))  '職等 
	Jstr=trim(rs("jstr"))
	Gstr=trim(rs("gstr"))
	zstr=trim(rs("zstr"))
	wstr=trim(rs("wstr"))
	ustr=trim(rs("ustr"))
	PERSONID=TRIM(RS("PERSONID"))	'身分証字號
	BHDAT=TRIM(RS("BHDAT"))	'簽約日(保險日)
	PASSPORTNO=TRIM(RS("PASSPORTNO"))	'護照號碼
	VISANO=TRIM(RS("VISANO"))	'簽證號碼
	PISSUEDATE=TRIM(RS("PISSUEDATE")) '護照簽發日
	PDUEDATE=TRIM(RS("PDUEDATE"))	'護照有效期
	VDUEDATE=TRIM(RS("VDUEDATE"))	'簽證有效期	
	PHONE=TRIM(RS("PHONE"))	'聯絡電話
	MOBILEPHONE=TRIM(RS("MOBILEPHONE"))	'手機
	HOMEADDR=TRIM(RS("HOMEADDR"))	'聯絡地址
	EMAIL=TRIM(RS("EMAIL"))	'EMAIL
	OUTDATe=TRIM(RS("OUTDATe"))	'離職日
	MEMO=TRIM(RS("MEMO"))	'其他說明
	GTDAT=TRIM(RS("GTDAT"))	'加入工團(年月)
	MARRYED = RS("MARRYED")    '婚姻狀況
	SCHOOL=RS("SCHOOL") '教育程度
	SHIFT=RS("SHIFT") '班別
	grps = rs("grps") 
	studyjob=RS("studyjob") '職能學習 
	
	'PHOTOS=TRIM(RS("PHOTOS"))	'照片檔名 
	PHOTOS=RS("EMPID")&".JPG"
	'-----------------------------------------
	'PHU = RS("PHU")
	'NN = RS("NN")
	'KT = RS("KT")
	'TTKH = RS("TTKH")
	'MT = RS("MT")
	'BB = RS("BB")
	
	BANKID = RS("BANKID") 
	
	wkd_no = RS("wkd_no") 
	wkd_duedate = RS("wkd_duedate") 
	experience = RS("experience")
	urgent_person = RS("urgent_person")
	releation = RS("releation")
	urgent_addr = RS("urgent_addr")
	urgent_tel = RS("urgent_tel")
	urgent_mobile=RS("urgent_mobile")
	bh_person=RS("bh_person")
	bh_personID=RS("bh_personID")
	
	
END IF
rs.close
SET RS=NOTHING 
conn.close
set conn=nothing


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">

<style>
.tdp
{
    border-bottom: 1 solid #000000;
    border-left:  1 solid #000000;
    border-right:  0 solid #ffffff;
    border-top: 0 solid #ffffff;
    font-size:12px;
}
.tabp
{
    border-color: #000000 #000000 #000000 #000000; 
    border-style: solid; 
    border-top-width: 2px; 
    border-right-width: 2px; 
    border-bottom-width: 2px; 
    border-left-width: 2px;
    font-size:12px;
}
</style>
</HEAD> 
<body topmargin=0 onbeforeprint="printsub.style.display='none';closeN.style.display='none';" onafterprint="printsub.style.display='';closeN.style.display='';">     
<TABLE WIDTH=580 CLASS=txt BORDER=0  cellspacing="0" cellpadding="0" ALIGN=CENTER>
	<tr height=35>
		<td align=right >
			<input border="0" src="../Picture/pic_print.gif" name="printsub" width="24" height="23" type="image"  onclick="self.print();">　
			<input border="0" src="../Picture/pic_close.gif" name="closeN" width="24" height="23" type="image"  onclick="vbscript:window.close()">
		</td>
	</tr>
<table>
<TABLE WIDTH=600 CLASS=txt BORDER=0  cellspacing="0" cellpadding="0" ALIGN=CENTER>
	<TR>
		<TD WIDTH=100>
			<img src="pic/<%=EMPID%>.jpg"  border=1 WIDTH=100 HEIGHT=130>
		</TD>
		<TD VALIGN=BOTTOM ALIGN=CENTER>
			<font class=txtBK></font><br>
			<font class=txtBK>越籍員工個人資料</FONT>
			
			<br><BR><BR>
			<TABLE WIDTH=490 CLASS=TXT BORDER=0>
				<TR height=20>
					<TD WIDTH=50 ALIGN=RIGHT>Số thẻ：</TD>
					<TD WIDTH=60  ><%=EMPID%></TD>
					<TD WIDTH=80 ALIGN=RIGHT>Quốc tịch：</TD>
					<TD WIDTH=80><%=COUNTRYDESC%></TD>
					<TD WIDTH=80 ALIGN=RIGHT>NVX：</TD>
					<TD ><%=INDAT%></TD>
				</TR> 
				<TR height=20>
					<TD ALIGN=RIGHT>Xuong：</TD>
					<TD><%=WSTR%></TD>
					<TD ALIGN=RIGHT>Đơn vị：</TD>
					<TD><%=GSTR%>-<%=ZSTR%></TD>
					<TD ALIGN=RIGHT>Chuc Vu：</TD>
					<TD><%=jstr%></TD>
				</TR>	 
			</TABLE>	
		</TD>
	</TR> 	
</TABLE>
<hr size=0	style='border: 1px dotted #999999;' align=center width=600>   
<TABLE WIDTH=600 CLASS="tabp" BORDER=0  cellspacing="0" cellpadding="4" ALIGN=CENTER> 
	<tr HEIGHT=30>
		<td width=100 align=right class="tdp" >Họ Tên：</td>
		<td width=170 class="tdp"><%=empnam_cn%> <%=empnam_vn%></td>
		<td width=90 align=right class="tdp" >Ngày Sinh：</td>             
		<td class="tdp"><%=bdy%></td>
	</tr> 
	<tr  HEIGHT=30>
		<td  align=right class="tdp" >Giới Tính：</td>
		<td class="tdp" >
			<%if sex="M" then%><img src="../picture/y01.gif" align="absmiddle" width=12 height=12>&nbsp;Nam <%else%><img src="../picture/n01.gif" align="absmiddle" width=12 height=12>&nbsp;Nam<%end if %>&nbsp;		
			<%if sex="F" then%><img src="../picture/y01.gif" align="absmiddle" width=12 height=12>&nbsp;Nữ <%else%><img src="../picture/n01.gif" align="absmiddle" width=12 height=12>&nbsp;Nữ<%end if %>&nbsp;		
		</td>
		<td  align=right class="tdp" >Hôn Nhân：</td>             
		<td class="tdp" >
			<%if marryed="Y" then%><img src="../picture/y01.gif" align="absmiddle" width=12 height=12>&nbsp;Da KH<%else%><img src="../picture/n01.gif" align="absmiddle" width=12 height=12>&nbsp;Da KH<%end if%>&nbsp;
			<%if marryed="N" then%><img src="../picture/y01.gif" align="absmiddle" width=12 height=12>&nbsp;Chua KH<%else%><img src="../picture/n01.gif" align="absmiddle" width=12 height=12>&nbsp;Chua KH<%end if%>&nbsp;
			<%if (marryed<>"N" and marryed<>"Y" ) then%><img src="../picture/y01.gif" align="absmiddle" width=12 height=12>&nbsp;khác<%else%><img src="../picture/n01.gif" align="absmiddle" width=12 height=12>&nbsp;khác<%end if%>&nbsp;
			
		</td>
	</tr>		
	<tr  HEIGHT=30>
		<td   align=right class="tdp">SỐ CMND：</td>
		<td   class="tdp"><%=personid%>&nbsp</td>
		<td  align=right class="tdp">Ngày cấp：</td>             
		<td class="tdp"><%=passportNo%>&nbsp</td>
	</tr>	
	<tr  HEIGHT=30>
		<td  align=right class="tdp">Nơi cấp：</td>
		<td  class="tdp"><%=visano%>&nbsp</td>
		<td  align=right class="tdp">Học lực：</td>
		<td   class="tdp"><%=school%>&nbsp;</td>
	</tr>		
	<tr  HEIGHT=30>
		<td  align=right class="tdp">ĐTDD：</td>             
		<td class="tdp"><%=mobilephone%>&nbsp;</td>
		<td  align=right class="tdp">E-Mail：</td>             
		<td class="tdp"><%=EMAIL%>&nbsp;</td>
	</tr>			
	<tr  HEIGHT=30>
		<td  align=right class="tdp">Địa chi：</td>
		<td  class="tdp" colspan=3><%=HOMEADDR%>&nbsp;</td>
	</tr>		
	<tr  HEIGHT=30>
		<td  align=right class="tdp">SO Tai Khoan：</td>
		<td   class="tdp" colspan=3><%=bankID%>&nbsp;</td>		
	</tr>	
	<tr HEIGHT=30>
		<td  align=right class="tdp" valign=top>Ghi Chú：<br><br><br><br><br><br></td>
		<td   class="tdp" colspan=3 valign=top><%=memo%>&nbsp;</td>		
	</tr>												
</table> 
<BR>
 
 
</body>
</html>
		
 
