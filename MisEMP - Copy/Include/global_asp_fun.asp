<%
function QueryFun(SqlStr,QArray)
	on error resume next
	err.Clear ()
	dim rs,iRows,jCols
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation =3
	rs.Open SqlStr,conn,1,1

	if rs.EOF =true then
		QueryFun=false
	else
		redim QArray(rs.RecordCount-1,rs.Fields.count-1)
		for iRows=0 to rs.RecordCount -1
			for jCols=0 to rs.Fields.count-1
				QArray(iRows,jCols)=rs(jCols)
			next
			rs.MoveNext 
		next
		QueryFun=true
	end if
	rs.Close
	set rs=nothing
	
	if err.number <> 0 then
		response.Write "SqlStr=" & SqlStr & "<br>"
		response.Write "err.number=" & err.number & "<br>"
		response.Write "err.Description=" & err.Description  & "<br>"
		response.Write "err.Source=" & err.Source  & "<br>"
		response.End 
	end if
end function

'--------------error found 回上頁function
function errortoback(msg)
	Response.Write "<script language=vbs>"
	Response.Write "alert (""" & msg &""")" & chr(10)+chr(13)
	Response.Write "window.history.go(-1)" & chr(10)+chr(13)
	Response.Write "</script>"
end function

'----------------字串補足"0"function
function getzero(num,length)
	dim len_num,i,zero_str,str
	len_num=len(num)
	    for i=1 to length-len_num
			zero_str="0"+zero_str
		next
	str=zero_str+cstr(num)
	getzero=str
end function

'------------------分頁function
'size		=分頁筆數
'sql		=sql語句
'page		=欲跳之頁數
'pagecount	=return 總共頁數
'recordcount=retrun 總筆數
'Gatedata	=return data
function AbsPage(size,sql,page,pagecount,recordcount,Getdata)
	on error resume next
	err.Clear ()
	
	dim rs
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation =3
	rs.Open sql,conn,1,1
	if not rs.EOF then
		rs.PageSize = size
		select case cint(page)
			case rs.PageCount,rs.PageCount+1
				page=rs.PageCount
				if rs.RecordCount mod size=0 then
					redim Gdata(size-1,rs.Fields.count-1)
				else
					redim Gdata((rs.RecordCount mod size)-1,rs.Fields.count-1)
				end if
			case 0
				page=1
				if cint(rs.RecordCount) < cint(size) then
					redim Gdata(rs.RecordCount -1,rs.Fields.count-1)
				else
					redim Gdata(size-1,rs.Fields.count-1)
				end if	
			case else
				page=cint(page)
				if cint(rs.RecordCount) < cint(size) then
					redim Gdata(rs.RecordCount -1,rs.Fields.count-1)
				else
					redim Gdata(size-1,rs.Fields.count-1)
				end if
		end select
		rs.AbsolutePage =page
		dim i,j
		i=0
		j=0
		for i=0 to rs.PageSize -1
			if rs.EOF then exit for
			for j=0 to rs.Fields.count-1
				Gdata(i,j)=rs.Fields(j)
				'Response.Write "gdata("&i&","&j&")="&Gdata(i,j)&"<br>"
			next
			rs.MoveNext 	
		next
		AbsPage	   =true
		Getdata    =Gdata
		pagecount  =rs.PageCount 
		recordcount=rs.RecordCount 
	else
		AbsPage=false
	end if
	rs.Close 
	set rs=nothing
	
	if err.number <> 0 then
		response.Write "sql=" & sql & "<br>"
		response.Write "err.number=" & err.number & "<br>"
		response.Write "err.Description=" & err.Description  & "<br>"
		response.Write "err.Source=" & err.Source  & "<br>"
		response.End 
	end if
end function

FUNCTION TRANS(STR)'-----數字轉為貨幣值顯示

DIM TEMP,INTSTR,CLEN,TEMP1,TEMP3,R,L
TEMP=""
INTSTR=""
CLEN=0
TEMP3=""
TEMP1=""
R=0
L=0

 IF TRIM(STR)="" OR TRIM(STR)="0" THEN
     TRANS_CURR="0"
 ELSE    
 
	 TEMP=SPLIT(STR,".")
	IF UBOUND(TEMP)>0 THEN
	 INTSTR=TRIM(REPLACE(TEMP(0),",",""))
	ELSE
	 INTSTR=TRIM(REPLACE(STR,",",""))
	END IF

	TEMP1=INTSTR'---右取整數(變動)
	TEMP3=INTSTR'---左取整數(固定)
	CLEN=LEN(INTSTR)'---原數字串長度
	'ALERT CLEN/3
	 K=0
		FOR I=1 TO CLEN/3
		
		  R=I*3+K
		  IF R<CLEN+k THEN'----當所取長度超過原字串長則停止
			TEMP1=","&RIGHT(TEMP1,R)'---每次由右邊取三位並在前面增加逗號
			'alert temp1
			'---------------------------
			L=CINT(CLEN-I*3)
			TEMP1=LEFT(TEMP3,L)&TEMP1'----每次由左邊取原字串長減去右邊所取字串數
			'alert temp1
			'---------------------------
		  else
			'alert r&"=="&clen
		  END IF 
			'ALERT TEMP1
		  K=K+1
		NEXT
		
		if ubound(TEMP)>0 THEN'----如有小數點則整數串與小數串相接
			TRANS=TEMP1&"."&TEMP(1)
		ELSE
			TRANS=TEMP1
		END IF
	'ALERT TEMP1
    	
  END IF  	
END FUNCTION
'--------------------使用者資料與所屬單位別選擇
FUNCTION SELECT_USERINFO(USERID,DATAarray)
	
	 Sql="EXEC SP_SYSUSER_INFO '"&SESSION("USERID")&"'"
	 INFOFlag=QueryFun(Sql,DATAarray)
	 
	 IF DATAarray(0,4)="PPPP" THEN
		 Sql="SELECT TBLCD,TBLDESC FROM YZZDCODE  WHERE TBLID='UNITNO'"
		 SFlag=QueryFun(Sql,Sarray)
		 	Response.Write "<SELECT NAME='USERUNIT' ID='USERUNIT'>"
		 FOR i=0 to UBOUND(Sarray)
		    Response.Write "<option value='"&Sarray(i,0)&"'>"&Sarray(i,0)&"-"&Sarray(i,1)&"</option>"
		 NEXT 
		    Response.Write "</SELECT>"
	 ELSE
		    Response.Write "使用者單位:<INPUT CLASS='READONLY' NAME='USERUNIT' ID='USERUNIT' VALUE='"&DATAarray(0,4)&"' READONLY>"	
	 END IF	        	 
	
END FUNCTION
'--------------------單位換算(原單位,新單位,比率1,比率2,大單位,中單位,小單位,原單位價格)
FUNCTION PUN_CHG(OPUN,NPUN,WP,PD,WN,PN,DN,NUM)
 
        OP=CINT(PUN_INDEX(OPUN,WN,PN,DN))
        NP=CINT(PUN_INDEX(NPUN,WN,PN,DN))

     if WN<>"" and PN<>"" and DN<>"" and cdbl(PD)<>0 then
				  If OP > NP then
				   
				   if  (OP - NP )=2 then'---小->大
				         PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)/cdbl(WP)/cdbl(PD),4))
				         'RESPONSE.WRITE "小->大 NUM="&PUN_CHG
					 else 
					      if OP<>1 then'---小->中
					       PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)/cdbl(PD),4))
					     ' RESPONSE.WRITE "小->中 NUM="&PUN_CHG
					      else
						   PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)/cdbl(WP),4))'----中->大
						  ' RESPONSE.WRITE "中->大 NUM="&PUN_CHG
						  end if 
					 end if
					 
				 elseif OP < NP then
				  
					 if  (NP - OP)=2 then'--大->小
				         PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)*cdbl(WP)*cdbl(PD),4))
				         'RESPONSE.WRITE "大->小 NUM="&PUN_CHG
					 else 
					      if OP=0 then'---大->中
					       PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)*cdbl(WP),4))
					      ' RESPONSE.WRITE "大->中 NUM="&PUN_CHG
					      else
						   PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)*cdbl(PD),4))'----中->小
						 ' RESPONSE.WRITE "中->小 NUM="&PUN_CHG
						  end if 
					 end if
				else
				   PUN_CHG=cdbl(Num)
				end if	 
	 else
	     IF PD=0 THEN PD=1 END IF'---舊原紙產品
	     
	     If OP > NP then
		         PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)/cdbl(WP)/cdbl(PD),4))
		         'RESPONSE.WRITE "小->大 NUM="&PUN_CHG
			 
		 elseif  OP < NP then
		         PUN_CHG=CDBL(FORMATNUMBER(cdbl(Num)*cdbl(WP)*cdbl(PD),4))
	            ' RESPONSE.WRITE "大->小 NUM="&PUN_CHG
		 else
		   PUN_CHG=cdbl(Num)
		 end if	
		  
	 end if 		 
              
              
END FUNCTION

'--------------單位順序轉換為數字順序
FUNCTION PUN_INDEX(PUN,WN,PN,DN)
   IF NOT ISNUMERIC(PUN) THEN
        IF WN<>"" AND PN<>"" AND DN<>"" THEN
			SELECT CASE PUN
			   CASE WN
			      PUN_INDEX=0
			   CASE PN
			      PUN_INDEX=1
			   CASE DN
			      PUN_INDEX=2        
			END SELECT
		ELSEIF WN<>"" THEN
		    SELECT CASE PUN
			   CASE PN
			      PUN_INDEX=0
			   CASE DN
			      PUN_INDEX=1        
			END SELECT
		ELSE
		   	SELECT CASE PUN
			   CASE WN
			      PUN_INDEX=0
			   CASE PN
			      PUN_INDEX=1
			   CASE DN
			      PUN_INDEX=1        
			END SELECT
		END IF		
   ELSE
        PUN_INDEX=PUN-1  
   END IF      
END FUNCTION
%>

<script language="vbscript">

FUNCTION TRANS_CURR1(STR)'-----數字轉為貨幣值顯示

DIM TEMP,INTSTR,CLEN,TEMP1,TEMP2,R,L
TEMP=""
INTSTR=""
CLEN=0
TEMP2=""
TEMP1=""
R=0
L=0
STR=cstr(STR)
 IF TRIM(STR)="" OR TRIM(STR)="0" THEN
     TRANS_CURR1="0"
 ELSE    
 
	 TEMP=SPLIT(STR,".")
	IF UBOUND(TEMP)>0 THEN
	 INTSTR=CDBL(REPLACE(TEMP(0),",",""))
	ELSE
	 INTSTR=CDBL(REPLACE(STR,",",""))
	END IF

	TEMP1=INTSTR'---右取整數(變動)
	TEMP2=INTSTR'---左取整數(固定)
	CLEN=LEN(INTSTR)'---原數字串長度
	'ALERT CLEN/3
	 K=0
		FOR I=1 TO CLEN/3
		
		  R=I*3+K
		  IF R<CLEN+k THEN'----當所取長度超過原字串長則停止
			TEMP1=","&RIGHT(TEMP1,R)'---每次由右邊取三位並在前面增加逗號
			'alert temp1
			'---------------------------
			L=CINT(CLEN-I*3)
			TEMP1=LEFT(TEMP2,L)&TEMP1'----每次由左邊取原字串長減去右邊所取字串數
			'alert temp1
			'---------------------------
		  else
			'alert r&"=="&clen
		  END IF 
			'ALERT TEMP1
		  K=K+1
		NEXT
		
		if ubound(TEMP)>0 THEN'----如有小數點則整數串與小數串相接
			TRANS_CURR1=TEMP1&"."&TEMP(1)
		ELSE
			TRANS_CURR1=TEMP1
		END IF
	'ALERT TEMP1
    	
  END IF  	
END FUNCTION

'--------------------單位換算(原單位,新單位,比率1,比率2,大單位,中單位,小單位,原單位價格)(SCRIPT)
FUNCTION SPUN_CHG(OPUN,NPUN,INWP,INPD,WNAME,PNAME,DNAME,NUM)

        OP=CINT(SPUN_INDEX(OPUN,WNAME,PNAME,DNAME))
        NP=CINT(SPUN_INDEX(NPUN,WNAME,PNAME,DNAME))
       
		if WNAME<>"" and PNAME<>"" and DNAME<>"" AND INPD<>0 then
				  If OP > NP then
				   
				   if  (OP - NP )=2 then'---小->大
				         SPUN_CHG=cdbl(Num)/cdbl(INWP)/cdbl(INPD)
				         'alert "小->大 NUM="&SPUN_CHG
					 else 
					      if OP<>1 then'---小->中
					       SPUN_CHG=cdbl(Num)/cdbl(INPD)
					      ' alert "小->中 NUM="&SPUN_CHG
					      else
						   SPUN_CHG=cdbl(Num)/cdbl(INWP)'----中->大
						   'alert "中->大 NUM="&SPUN_CHG
						  end if 
					 end if
					 
				 elseif cint(OP) < NP then
				  
					 if  (NP - OP)=2 then'--大->小
				         SPUN_CHG=cdbl(Num)*cdbl(INWP)*cdbl(INPD)
				         'alert "大->小 NUM="&SPUN_CHG
					 else 
					      if OP=0 then'---大->中
					       SPUN_CHG=cdbl(Num)*cdbl(INWP)
					       'alert "大->中 NUM="&SPUN_CHG
					      else
						   SPUN_CHG=cdbl(Num)*cdbl(INPD)'----中->小
						  ' alert "中->小 NUM="&SPUN_CHG
						  end if 
					 end if
				else
				   SPUN_CHG=cdbl(Num)
				end if	 
		else
		
		    if INPD=0 then INPD=1 end if'---舊原紙類inpd值為零
		    
		    If OP > NP then
			         SPUN_CHG=cdbl(Num)/cdbl(INWP)/cdbl(INPD)
			         'alert "小->大 NUM="&SPUN_CHG
				 
			 elseif cint(OP) < NP then
			         SPUN_CHG=cdbl(Num)*cdbl(INWP)*cdbl(INPD)
		            'alert "大->小 NUM="&SPUN_CHG
			 else
			   SPUN_CHG=cdbl(Num)
			 end if	
			  
		end if 		 
              
              
END FUNCTION

'--------------單位順序轉換為數字順序(SCRIPT)
FUNCTION SPUN_INDEX(PUN,WNAME,PNAME,DNAME)
   IF NOT ISNUMERIC(PUN) THEN
        IF WNAME<>"" AND PNAME<>"" AND DNAME<>"" THEN
			SELECT CASE PUN
			   CASE WNAME
			      SPUN_INDEX=0
			   CASE PNAME
			      SPUN_INDEX=1
			   CASE DNAME
			      SPUN_INDEX=2        
			END SELECT
		ELSEIF WNAME<>"" THEN
		    SELECT CASE PUN
			   CASE PNAME
			      SPUN_INDEX=0
			   CASE DNAME
			      SPUN_INDEX=1        
			END SELECT
		ELSE
		   	SELECT CASE PUN
			   CASE WNAME
			      SPUN_INDEX=0
			   CASE PNAME
			      SPUN_INDEX=1
			   CASE DNAME
			      SPUN_INDEX=1        
			END SELECT
		END IF		
   ELSE
        SPUN_INDEX=PUN-1  
   END IF      
END FUNCTION
</script>