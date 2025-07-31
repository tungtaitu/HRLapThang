function chkdat(sid){	
	var ans = document.getElementById(sid).value ;  
	if (ans !="") {
		if (ans.length <8  || ans.length > 10 ) { 			
			//document.getElementById("hit").innerText="日期輸入錯誤" ;
			//document.getElementById("hit").style.color="red";			
			alert ("日期輸入錯誤yyyy.mm.dd");
			document.getElementById(sid).value = "" ;
			window.setTimeout( function(){document.getElementById(sid).focus(); }, 0);
			return ; 
		}
		else { 
			if (ans.length ==8)  { 
				var y1= (parseInt(ans.substring(0,4),10)) ;
				var m1= (parseInt(ans.substr(4,2),10)) ; 				
				var d1= (parseInt(ans.substring(ans.length-2,ans.length),10)) ; 				
				//alert ( y1+"\n"+ m1+"\n"+d1) ;
				isYdate(y1,m1,d1,sid) ; 
			}			
			else if (ans.length==10)  { 
				var y1= (parseInt(ans.substring(0,4),10)) ;
				var m1= (parseInt(ans.substr(5,2),10)) ; 				
				var d1= (parseInt(ans.substring(ans.length-2,ans.length),10)) ; 							
				isYdate(y1,m1,d1,sid) ;
			}			
			else {
				document.getElementById(sid).value = "" ;
				//document.getElementById("hit").innerText="日期輸入錯誤" ;
				//document.getElementById("hit").style.color="red";		
				alert ("日期輸入錯誤yyyy.mm.dd");
				document.getElementById(sid).focus();				
			}
		}	
	} 	
}	
function isYdate(arg_intYear,arg_intMonth,arg_intDay,sid)
{
  var objDate = new Date(arg_intYear,arg_intMonth-1,arg_intDay);
  //檢查月份是否小於12大於1
  if((parseInt(arg_intMonth) > 12) || (parseInt(arg_intMonth) < 1))
  {
    //alert(arg_intYear+'/'+arg_intMonth+'/'+arg_intDay+'  月份不正確');
		document.getElementById(sid).value=""
		//document.getElementById("hit").innerText="日期輸入錯誤" ;
		//document.getElementById("hit").style.color="red";
		alert ("日期輸入錯誤yyyy.mm.dd");
		document.getElementById(sid).focus();
  }
  else
  {
    //如果objDate日數進位不等於傳入的arg_intDay，代表天數格式錯誤，另外月份進位也代表日期格式錯誤
    if((parseInt(arg_intDay) != parseInt(objDate.getDate()))||(parseInt(arg_intMonth)!= parseInt((objDate.getMonth()+1))))
    {
      //alert(arg_intYear+'/'+arg_intMonth+'/'+arg_intDay+ '   天數不正確');
			document.getElementById(sid).value=""
			//document.getElementById("hit").innerText="日期輸入錯誤" ;
			//document.getElementById("hit").style.color="red";
			alert ("日期輸入錯誤yyyy.mm.dd");
			document.getElementById(sid).focus();
    }
    else
    {
      //alert(arg_intYear+'/'+arg_intMonth+'/'+arg_intDay+ '  日期格式正確');
			var mm="00"+String(arg_intMonth)  ;
			var dd="00"+String(arg_intDay)  ;
			//document.getElementById("hit").innerText="" ;
			document.getElementById(sid).value=arg_intYear+'.'+mm.substring(mm.length-2,mm.length)+'.'+dd.substring(dd.length-2,dd.length)  ;
    }
  }
}