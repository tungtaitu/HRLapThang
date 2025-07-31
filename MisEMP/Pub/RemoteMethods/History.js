/* show history list detail */
function showHistory(objSender){
    var trDetail = $(objSender).parent().parent().next().next();
    var tdDetail = $(objSender).parent().parent().next().next().find("td:eq(0)");    
    //loading data
    if(jQuery.trim($(tdDetail).html())==""){
        //var vou_no = $(objSender).parent().parent().find("span[id$='lblvou_no']").text();
        var vou_no = $(objSender).parent().parent().find("span")[0].innerHTML; //feng 20161012

        var submit_id = $(objSender).parent().parent().find("td").find(".hidSubmitID").val();        
        getHistory(tdDetail, vou_no, submit_id);        
    }
    //toggle power list detail show/hide
    if(jQuery.trim($(trDetail).css("display"))=="none"){
        $(objSender).parent().parent().next().hide(); 
        $(trDetail).show();
    }else{
        $(trDetail).hide();
    }
    //release ram        
    trDetail = null;
    tdDetail = null;
}


function getHistory(objSender, vou_no, submit_id) {
    var objServerVar;
    $(objSender).html("<div style='color:gray;text-align:center'>Loading data, please wait ....<br><img src='../../App_Themes/pwfBody/images/AjaxLoadingBar1.gif'></div>");
	$.ajax(
		{
			url:"../../Pub/RemoteMethods/History.aspx",
			type:"post",
			data: "EventName=getHistory&vou_no=" + vou_no + "&submit_id=" + submit_id,
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
					$(objSender).html(objServerVar.Result );
				}else{
				    alert(objServerVar.ServerMsg);
				}
			}
			
		}
	);        
}

