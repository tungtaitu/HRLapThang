<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
SELF = "vyfysuco"

func = request("func")
code = request("code")

Set conn = GetSQLServerConnection()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head>
<%
select case func
	case "A"
		sql="select * from view_empfile where empid = '"& code &"'  "
		response.write sql
		'response.end
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
  		if not rst.eof then
%>			<script language=vbs>
				parent.Fore.<%=self%>.empid.value="<%=code%>"
				parent.Fore.<%=self%>.whsno.value="<%=rsT("whsno")%>"
				parent.Fore.<%=self%>.cfdw.value="<%=rsT("empnam_CN")%>"
			</script>
<% 		else %>
			<script language=vbs>
				alert "員工輸入錯誤"
				parent.Fore.<%=self%>.empid.value=""
				parent.Fore.<%=self%>.whsno.value=""
				parent.Fore.<%=self%>.empid.focus()
			</script>
<%
		end if
		set rs=nothing
	case "B"
		sql="select * from yfymsuco where sgno='"& Trim(request("code1")) &"' and pddate='"& Trim(request("code2")) &"'  "
		set rds=conn.execute(sql)
		if not rds.eof then
%>			<script language=vbs>
				Parent.Fore.<%=self%>.autoid.value="<%=rds("autoid")%>"
				Parent.Fore.<%=self%>.TOTcost.value="<%=rDs("TOTcost")%>"
				Parent.Fore.<%=self%>.SGMEMO.value="<%=RDS("SGMEMO")%> "
				Parent.Fore.<%=self%>.cfgroup.focus()
			</script>
<% 		end if
		set rs=nothing

	case "CDATACHG"
		tmpRec(CurrentPage,index + 1,0) = "UPD"
  		tmpRec(CurrentPage,index + 1,23) = CODESTR01
  		tmpRec(CurrentPage,index + 1,24) = CODESTR02
  		tmpRec(CurrentPage,index + 1,25) = CODESTR03
  		tmpRec(CurrentPage,index + 1,26) = CODESTR04
  		tmpRec(CurrentPage,index + 1,27) = CODESTR05
  		tmpRec(CurrentPage,index + 1,32) = CODESTR06
  		tmpRec(CurrentPage,index + 1,35) = CODESTR07
  		tmpRec(CurrentPage,index + 1,37) = CODESTR08
  		tmpRec(CurrentPage,index + 1,58) = CODESTR09
  		tmpRec(CurrentPage,index + 1,34) = CODESTR10
  		tmpRec(CurrentPage,index + 1,36) = CODESTR11
  		TTM = cdbl(tmpRec(CurrentPage, index+1, 20))+cdbl(tmpRec(CurrentPage, index+1, 22))+cdbl(tmpRec(CurrentPage, index+1, 23))
  		if tmpRec(CurrentPage,index + 1,4)="VN" then
	  		TTMH = round( (TTM/26/8), 0 )
	  	else
	  		TTMH = round( (TTM/30/8), 3 )
	  	end if
  		tmpRec(CurrentPage,index + 1,38)=TTMH

  		'+(加項)
  		BB = tmpRec(CurrentPage,index + 1,20)
  		CV = tmpRec(CurrentPage,index + 1,22)
  		PHU = tmpRec(CurrentPage,index + 1,23)
  		NN = tmpRec(CurrentPage,index + 1,24)
  		KT = tmpRec(CurrentPage,index + 1,25)
  		MT = tmpRec(CurrentPage,index + 1,26)
  		TTKH = tmpRec(CurrentPage,index + 1,27)
  		QC = tmpRec(CurrentPage,index + 1,31)
  		TNKH = tmpRec(CurrentPage,index + 1,32)
  		TBTR = tmpRec(CurrentPage,index + 1,33)
  		JX = tmpRec(CurrentPage,index + 1,58)  '績效獎金

  		response.write   "BB=" & BB &"<BR>"
  		response.write   "CV=" & CV &"<BR>"
  		response.write   "PHU=" & PHU &"<BR>"
  		response.write   "NN=" & NN &"<BR>"
  		response.write   "KT=" & KT &"<BR>"
  		response.write   "MT=" & MT &"<BR>"
  		response.write   "TTKH=" & TTKH &"<BR>"
  		'-(減項)
  		BH = tmpRec(CurrentPage,index + 1,34)
  		HS = tmpRec(CurrentPage,index + 1,35)
  		GT = tmpRec(CurrentPage,index + 1,36)
  		QITA = tmpRec(CurrentPage,index + 1,37)


  		if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then
  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then
	  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC) )/26 ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
	  		else
	  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU) )/26 ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+CDBL(QC)
	  		end if
	  	else
	  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
	  	end if

	  	RESPONSE.WRITE 	"F1_MONEY=" & F1_MONEY &"<br>"
	  	tmpRec(CurrentPage,index + 1,60) = CDBL(F1_MONEY)

	  	F2_MONEY =CDBL(BH)+CDBL(HS)+CDBL(GT)+CDBL(QITA)    '應扣款
	  	RESPONSE.WRITE 	"F2_MONEY=" & F2_MONEY &"<br>"

	  	'tmpRec(CurrentPage,index + 1,50) : 曠職事假病假扣款
	  	'tmpRec(CurrentPage,index + 1,49) 總加班費

	  	TMONEY = CDBL(F1_MONEY) - CDBL(F2_MONEY)
	  	tmpRec(CurrentPage,index + 1,39) = TMONEY '應發金額     (不含加班扣減時假)
	  	RESPONSE.WRITE "TMONEY=" & TMONEY &"<br>"
	  	relTOTM = TMONEY + cdbl(tmpRec(CurrentPage,index + 1,49)) - cdbl(tmpRec(CurrentPage,index + 1,50))
  		tmpRec(CurrentPage,index + 1,47)  = relTOTM  '實領金額(含加班扣減時假)
  		tmpRec(CurrentPage,index + 1,64)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(tmpRec(CurrentPage,index + 1,49))
	  	tmpRec(CurrentPage,index + 1,65) = cdbl(tmpRec(CurrentPage,index + 1,64))-cdbl(F2_MONEY)-cdbl(tmpRec(CurrentPage,index + 1,50))-cdbl(tmpRec(CurrentPage,index + 1,47))

%>		<script language=vbs>
			Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=(TTMH)%>
			'Parent.Fore.<%=self%>.TOTMONEY(<%=index%>).value=<%=(TMONEY)%>
			'Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=(F1_MONEY)%>
			Parent.Fore.<%=self%>.BZKM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,65)%>
			Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,64)%>
			Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value=<%=(relTOTM)%>
		'	alert <%=TTMH%>
		</script>
<%
end  select
Session("empfilesalary") = tmpRec
%>
</html>
