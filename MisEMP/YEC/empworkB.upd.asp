<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  
<%
Response.Expires = 0
Response.Buffer = true 

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")
INDAT = request("INDAT")
Set conn = GetSQLServerConnection()
tmpRec = Session("empworkbC") 

Set CONN = GetSQLServerConnection()  

conn.BeginTrans
x = 0
y = ""  
empid = request("empid") 
'response.write empid &"<BR>"
YYMM=REQUEST("YYMM")
EMPIDSTR="" 
for i = 1 to TotalPage 
	for j = 1 to PageRec   
		
		 totJia = cdbl(tmpRec(i, j, 9))+cdbl(tmpRec(i, j, 10))+cdbl(tmpRec(i, j, 11))+cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 21))		 
		 tmpRec(i,j, 2) = request("TIMEUP")(j) 
		 tmpRec(i,j, 3) = request("TIMEDOWN")(j) 
		 tmpRec(i,j, 4) = request("TOTHOUR")(j) 
		 tmpRec(i,j, 5) = request("H1")(j) '平日加班
		 tmpRec(i,j, 6) = request("H2")(j) '休息加班
		 tmpRec(i,j, 7) = request("H3")(j) '假日加班
		 tmpRec(i,j, 8) = request("B3")(j) '夜班
		 tmpRec(i,j,20) = request("Forget")(j)  '忘刷卡
		 tmpRec(i,j,19) = request("KZhour")(j)  '曠職
		 tmpRec(i,j,31) = request("LATEFOR")(j)  '遲到
		 'response.write tmpRec(i,j,1) &"<BR>" 
		 'response.write tmpRec(i,j,2) &"<BR>" 
		 'response.write tmpRec(i,j,3) &"<BR>" 
		 'response.write tmpRec(i,j,4) &"<BR>" 
		 'response.write tmpRec(i,j,5) &"<BR>" 
		 'response.write tmpRec(i,j,6) &"<BR>" 
		 'response.write tmpRec(i,j,7) &"<BR>" 
		 'response.write tmpRec(i,j,8) &"<BR>" 
		 'response.write tmpRec(i,j,20) &"<BR>" 
		 'response.write tmpRec(i,j,19) &"<BR>" 
		 'response.write j &"<BR>"    
		 
		 if tmpRec(i,j, 2)<>"" then 
		 	T1= replace(tmpRec(i,j, 2),":","") & "00"
		 else
		 	T1="000000" 
		 end if		 
		 if tmpRec(i,j, 3)<>"" then 
		 	T2= replace(tmpRec(i,j, 3),":","") & "00"
		 else
		 	T2="000000" 
		 end if	 
		 sql2="select * from empwork where empid='"& empid &"' and workdat='"& replace(tmpRec(i,j,1),"/","") &"'  "   
		 Set rs = Server.CreateObject("ADODB.Recordset")     
		 RS.OPEN SQL2, CONN, 3, 3  
		 IF RS.EOF THEN 		 	
		 	if  cdate(trim((tmpRec(i,j,1)))) >=cdate(indat) and cdate(trim((tmpRec(i,j,1))))  <= cdate(date())  then 
			 	SQL = "INSERT INTO empwork (EMPID , workdat, timeup, timedown, toth, forget, kzhour , H1, H2, H3, B3 , "&_
			 		  "latefor , flag , yymm, mdtm, muser , userIP   ) values ( "&_
			 		  "'"& empid &"', '"& replace(tmpRec(i,j,1),"/","") &"', '"& T1 &"', '"& T2 &"' , '"& tmpRec(i,j, 4) &"' ,"&_
			 		  "'"& cdbl(tmpRec(i,j,20)) &"','"& tmpRec(i,j,19) &"', '"& tmpRec(i,j,5) &"', '"& tmpRec(i,j,6) &"', "&_
			 		  "'"& tmpRec(i,j,7) &"', '"& tmpRec(i,j,8) &"', '"& cdbl(tmpRec(i,j,31)) &"', 'UPD' , '"& YYMM &"', getdate(), "&_
			 		  "'"& session("netuser") &"', '"& session("vnLogIP") &"'   ) " 
			 	conn.execute(sql) 	  
			 	X = X + 1
			end if  	
		 ELSE 
		    if cdate(trim((tmpRec(i,j,1)))) <= cdate(date())  then 
    		 	if rs("flag")<>"JIA"  or totJia>0 then  				
                    sql="update empwork set timeup='"& T1 &"' , timedown='"& T2 &"' ,  "&_
    					 "ToTH='"& tmpRec(i,j, 4) &"', forget='"& cdbl(tmpRec(i,j,20)) &"'  , "&_
    				 	 "kzhour='"& tmpRec(i,j,19) &"', H1='"& tmpRec(i,j,5) &"'  , "&_
    				 	 "H2='"& tmpRec(i,j,6) &"', H3='"& tmpRec(i,j,7) &"'  , B3='"& tmpRec(i,j,8) &"',  "&_
    				 	 "latefor='"& cdbl(tmpRec(i,j,31)) &"', flag='UPD' ,  "&_
    				 	 "mdtm=getdate(), muser='"& session("netuser") &"', userip='"& session("vnLogIP") &"'   "&_
    				 	 "where empid='"& empid &"' and workdat='"& replace(tmpRec(i,j,1),"/","") &"' "
                    response.write sql &"<BR>"    
                    X = X + 1
                    conn.execute(sql)  
    			else 				
        			sql="update empwork set kzhour='"& tmpRec(i,j,19) &"' , "&_
        				"ToTH='"& tmpRec(i,j, 4) &"' , forget='"& cdbl(tmpRec(i,j,20)) &"' , latefor='"& cdbl(tmpRec(i,j,31)) &"' , "&_
        				"mdtm=getdate(), muser='"& session("netuser") &"', userip='"& session("vnLogIP") &"'   "&_
        				"where empid='"& empid &"' and workdat='"& replace(tmpRec(i,j,1),"/","") &"' "
        			response.write "b=" & sql &"<BR>" 
        			conn.execute(sql) 
    			end if 	 
    		end  if	
		END IF   
		set rs = nothing 
		 
	next
next 	

'RESPONSE.END 

if conn.Errors.Count = 0 or err.number=0  then 
	conn.CommitTrans
	'Set Session("empworkbC") = Nothing 	    
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		'ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans SUCCESS !!"
		'OPEN "empfile.salary.asp" , "_self" 
		window.close()
	</script>	
<%
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		'ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		'OPEN "empfile.salary.asp" , "_self" 
		window.close()
	</script>	
<%	response.end 
END IF  
%>
 