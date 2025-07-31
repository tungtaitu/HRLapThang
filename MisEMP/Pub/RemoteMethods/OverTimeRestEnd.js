
function ShowDetail(objSender,sdate,edate){
    var trDetail = $(objSender).parent().parent().next();
    var tdDetail = $(objSender).parent().parent().next().find("td:eq(0)");    
        
    if(jQuery.trim($(trDetail).css("display"))=="none"){
        //$(objSender).parent().parent().next().next().next().next().hide(); 
        $(trDetail).show();

        var emp_no = $(objSender).parent().parent().find("span[id$='ctrEmp_No']").text();
        getDetail(tdDetail, emp_no, sdate, edate);        
    }else{
        $(trDetail).hide();
    }
    trDetail = null;
    tdDetail = null;
}

function getDetail(objSender,emp_no,sdate,edate)
{
    
    var objServerVar;
    $(objSender).html("<div style='color:gray;text-align:center'>Loading data, please wait ....<br><img src='../../../App_Themes/pwfBody/images/AjaxLoadingBar1.gif'></div>");
	$.ajax(
		{
			url:"../../../Pub/RemoteMethods/OverTimeRestEnd.aspx",
			type:"post",
			data:"EventName=OverTimeRestEnd&emp_no="+emp_no+"&start_date="+sdate+"&end_date="+edate,
			dataType:"text",
			timeout:100000,
			error:function(err){
				objServerVar = {
					IsOK:false,
					ServerMsg : "遠端連線發生錯誤或連線逾時，請再試一次，如果持續發生請連絡相關人員！！"
				}
				$(objSender).html("");
				$(objSender).parent().hide();
			},
			success:function(strServerVar){		
				eval("objServerVar ="+strServerVar);
			},
			complete:function(){			
				if(objServerVar.IsOK){		    	
					$(objSender).html(objServerVar.Result  );
				}else{
				    alert(objServerVar.ServerMsg);
				}
			}
			
		}
	);        
}

