<%@Language=VBScript CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<!-------- #include file = "../../GetSQLServerConnection.fun" --------->
<!--#include file="../../ADOINC.inc"-->
<%

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage") 

xhid=request("xhid")
lorry=request("lorry")
soxe=request("soxe")
ton = request("ton")
xeindat = request("xeindat")
if xeindat="" then 
	xeindat="null"
else
	xeindat="'"&xeindat&"'"	
end if 

whsno = request("whsno")
wbloai = request("wbloai")


Set conn = GetSQLServerConnection()
tmpRec = Session("YEBE0103B")
'on error resume next 
if session("netuser")="" then 
	response.write "使用者帳號為空,請重新登入!!<BR>"
	response.write "UserID is Null Please Login again !!!<BR>"
	response.write "Vao mang trong rong , hoac doi lau , hay nhan nut nhap mang tu dau !!! "
	response.end 
end if 
 	
 	
conn.BeginTrans 

response.write gTotalPage &"<BR>" 
response.write PageRec &"<BR>"
'response.end 

x = 0
y = ""
for i = 1 to pagerec 	
	op=trim(request("op")(i))
	wbid=trim(request("wbid")(i))
	cv=trim(request("cv")(i))
	name_vn=trim(request("name_vn")(i))
	txindat=trim(request("txindat")(i))
	dt1=trim(request("dt1")(i))
	dt2=trim(request("dt2")(i))
	aid=trim(request("aid")(i))
	xhid=trim(request("xhid")(i))
	nlorry=trim(request("nlorry")(i))
	nsoxe=trim(request("nsoxe")(i))
	
	if op="Y" then 
		if trim(wbid)<>"" then 
			if txindat="" then 
				txindat="null"
			else
				txindat="'"&txindat&"'"
			end if	
			
			sql="insert into wbempfile (wbwhsno,loai, wbid, wbname_vn, indat, fac, lorry, soxe , job, mdtm, muser, phone, mobile ) values( "&_
				"'"&whsno&"','"&wbloai&"','"&wbid&"',N'"&name_vn&"',"&txindat&",'"&xhid&"', "&_
				"'"&nlorry&"','"&nsoxe&"','"&cv&"',getdate(),'"&session("netuser")&"','"&DT1&"','"&DT2&"')  " 
			response.write sql &"<BR>"	
			conn.execute(sql)
			if aid<>"" then 
				sql2="update [yfymis].dbo.ysbdxetp set wbid='"&wbid&"' where aid='"&aid&"'"
				conn.execute(sql2)
			end if 	
		end if 	
	end if  
next 
'RESPONSE.WRITE Y 
'response.end  

 if conn.Errors.Count = 0 then 
	conn.CommitTrans
	Set Session("YEBE0103B") = Nothing    
	conn.close
	Set conn = Nothing

	Session("Title") = ""
	Session("Name") = "IMessage"
	Session("NO") = "Data Complete Success (OK) Count:" & x 
	Session("MessageCode") = "Success"
	Session("KeyValue") =  Y
	Session("SubmitValue") = replace(session("pgname"),"<BR>",chr(13))
	'Session("Action") = "YSBAE1403.Fore.asp"
	Response.Redirect "YEBE0104.Fore.asp?whsno="&whsno&"&wbloai="&wbloai&"&soxe="&soxe
	Set conn = Nothing 
 else
	conn.RollbackTrans 
    Set Session("YEBE0103B") = Nothing
    Set cmd = Nothing
	conn.close	
	Set conn = Nothing 
%>	<script language=vbs>
		alert "資料處理失敗(Fial)!!"
		open "yebe0103.asp"
	</script>
	
<% end if  %>
