<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
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
			if tmpRec(i,j,20)="" then tmpRec(i,j,20) = 0 
			tmpRec(i,j,19) = request("KZhour")(j)  '曠職
			if tmpRec(i,j,19) = "" then tmpRec(i,j,19) = 0 
			tmpRec(i,j,25) = request("LATEFOR")(j)  '遲到 
			if tmpRec(i,j,25)="" then tmpRec(i,j,25)= 0 
				
			tmpt1 	=  request("tmpt1")(j)  '修息時間1
			tmpt2 	=  request("tmpt2")(j)  '歇息時間2
			
			if ( tmpt1<>"" and len(tmpt1)>1 ) or ( tmpt2<>"" and len(tmpt2)>1 ) then 
				sqlx="if not exists (select * from empwork_xx where emp_id='"&empid&"' and work_dat='"&replace(tmpRec(i,j,1),"/","")&"' ) "&_
						 "insert into  empwork_xx ( emp_id,work_dat, tmpt1, tmpt2, mdtm,muser ) values ( '"&empid&"','"&replace(tmpRec(i,j,1),"/","")&"'  "&_
						 ",'"&tmpt1&"','"&tmpt2&"',getdate(),'"& session("netuser")&"')   "&_
						 "else "&_
						 "update  empwork_xx set tmpt1='"&tmpt1&"' , tmpt2='"&tmpt2&"', mdtm=getdate(), muser='"& session("netuser") &"' "&_
						 "where  emp_id='"&empid&"' and work_dat='"&replace(tmpRec(i,j,1),"/","")&"'  "  
				conn.execute(Sqlx)		 
				'response.write sqlx &"<BR>"		 
			elseif tmpt1="" and   tmpt2="" then 
				sql="delete empwork_xx where emp_id='"&empid&"' and work_dat='"&replace(tmpRec(i,j,1),"/","")&"' " 
				conn.execute(Sql)				
			end if 
			'response.write "a="&trim(request("flag")(j))&"<br>"
			if trim(request("flag")(j))="Y" then  	
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
			
				if yymm>="200810"  then 
					if right(tmpRec(i,j,1),2)>="26" then 
						if right(yymm,2)+1>13 then 
							yymm=left(yymm,4)+1&"01"
						else
							yymm=yymm+1 
						end if 				
					end if 
				end if 
				'response.write yymm 
				'response.end
				sql2="select * from empwork where empid='"& empid &"' and workdat='"& replace(tmpRec(i,j,1),"/","") &"'  "   
				Set rs = Server.CreateObject("ADODB.Recordset")     
				RS.OPEN SQL2, CONN, 3, 3  
				IF RS.EOF THEN 	 
					response.write "111111" 
					if  cdate(trim((tmpRec(i,j,1)))) >=cdate(indat) and cdate(trim((tmpRec(i,j,1))))  <= cdate(date())  then 
						SQL = "INSERT INTO empwork (EMPID , workdat, timeup, timedown, toth, forget, kzhour , H1, H2, H3, B3 , "&_
							  "latefor , flag , yymm, mdtm, muser , userIP , newtoth  ) values ( "&_
							  "'"& empid &"', '"& replace(tmpRec(i,j,1),"/","") &"', '"& T1 &"', '"& T2 &"' , '"& tmpRec(i,j, 4) &"' ,"&_
							  "'"& cdbl(tmpRec(i,j,20)) &"','"& tmpRec(i,j,19) &"', '"& tmpRec(i,j,5) &"', '"& tmpRec(i,j,6) &"', "&_
							  "'"& tmpRec(i,j,7) &"', '"& tmpRec(i,j,8) &"', '"& cdbl(tmpRec(i,j,25)) &"', 'UPD' , '"& YYMM &"', getdate(), "&_
							  "'"& session("netuser") &"', '"& session("vnLogIP") &"' ,'"& tmpRec(i,j, 4) &"' ) " 
						conn.execute(sql) 
						response.write sql &"<BR>"
						X = X + 1
					end if  	
				ELSE 
					response.write "2222222" 
					if cdate(trim((tmpRec(i,j,1)))) <= cdate(date())  then 
						if rs("flag")<>"JIA"  or totJia>0 then  				
							sql="update empwork set timeup='"& T1 &"' , timedown='"& T2 &"' ,  "&_
    						 "ToTH='"& tmpRec(i,j, 4) &"', forget='"& (tmpRec(i,j,20)) &"'  , "&_
    					 	 "kzhour='"& tmpRec(i,j,19) &"', H1='"& tmpRec(i,j,5) &"'  , "&_
    					 	 "H2='"& tmpRec(i,j,6) &"', H3='"& tmpRec(i,j,7) &"'  , B3='"& tmpRec(i,j,8) &"',  "&_
    					 	 "latefor='"&(tmpRec(i,j,25)) &"', flag='UPD' ,  "&_
    					 	 "mdtm=getdate(), muser='"& session("netuser") &"', userip='"& session("vnLogIP") &"' ,  "&_
    					 	 "newtoth='"&tmpRec(i,j,4)&"' where empid='"& empid &"' and workdat='"& replace(tmpRec(i,j,1),"/","") &"' "
							response.write sql &"<BR>"    
							X = X + 1                    	
							conn.execute(sql)   
							'response.end 			
						else  
							sql="update empwork set timeup='"& T1 &"' , timedown='"& T2 &"' ,  "&_
	    						 "ToTH='"& tmpRec(i,j, 4) &"', forget='"& cdbl(tmpRec(i,j,20)) &"'  , "&_
	    					 	 "kzhour='"& tmpRec(i,j,19) &"', H1='"& tmpRec(i,j,5) &"'  , "&_
	    					 	 "H2='"& tmpRec(i,j,6) &"', H3='"& tmpRec(i,j,7) &"'  , B3='"& tmpRec(i,j,8) &"',  "&_
	    					 	 "latefor='"& cdbl(tmpRec(i,j,25)) &"', flag='UPD' ,  "&_
	    					 	 "mdtm=getdate(), muser='"& session("netuser") &"', userip='"& session("vnLogIP") &"'   "&_
	    					 	 ",newtoth='"&tmpRec(i,j,4)&"' where empid='"& empid &"' and workdat='"& replace(tmpRec(i,j,1),"/","") &"' "
							response.write sql &"<BR>"    
							X = X + 1
							conn.execute(sql)  	        		
	        			end if	 
					end if 	 
				end  if				   
				set rs = nothing 
			End if 		 
		next
	next 	

'RESPONSE.END 

if conn.Errors.Count = 0 or err.number=0 then 
	conn.CommitTrans
	'Set Session("empworkbC") = Nothing 	    
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans SUCCESS !!"
		'OPEN "empfile.salary.asp" , "_self" 
		window.close()
	</script>	
<%
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		'OPEN "empfile.salary.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 