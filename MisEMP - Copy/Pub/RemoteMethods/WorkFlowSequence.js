/* show power list detail */
function showPowerDetail(objSender){
    var trDetail = $(objSender).parent().parent().next();
    var tdDetail = $(objSender).parent().parent().next().find("td:eq(0)");
    //loading data
    if(jQuery.trim($(tdDetail).html())==""){
        var strUserID = $(objSender).parent().find(":hidden:eq(0)").val();
        getPowerDetail(tdDetail,strUserID);
    }
    //toggle power list detail show/hide
    if(jQuery.trim($(trDetail).css("display"))=="none"){
        $(trDetail).show();
    }else{
        $(trDetail).hide();
    }
    //release ram        
    trDetail = null;
    tdDetail = null;
}

function getPowerDetail(objSender,strUserID){
    var objServerVar;
    $(objSender).html("<div style='color:gray;text-align:center'>Loading data, please wait ....<br><img src='../../../App_Themes/pwfBody/images/AjaxLoadingBar1.gif'><br></div><br>");
	$.ajax(
		{
			url:"../../../Pub/RemoteMethods/WorkFlowSequence.aspx",
			type:"post",
			data:"EventName=getPowerDetail&UserID="+strUserID,
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
					$(objSender).html(  objServerVar.Result );
				}else{
				    alert(objServerVar.ServerMsg);
				}
			}
			
		}
	);        
}

function delUserAuth(objSender) {    
    var strObjName = $(objSender).parent().parent().find("span")[1].innerHTML;
		strObjName += "-";
		strObjName += $(objSender).parent().parent().find("span")[2].innerHTML;
		strObjName += "-";
		strObjName += $(objSender).parent().parent().find("span")[3].innerHTML;

    // feng 20170321 
    /*var strObjName = $(objSender).parent().parent().find("span[id$='Label2']").text();
		strObjName += "-";
		strObjName += $(objSender).parent().parent().find("span[id$='Label3']").text();
		strObjName += "-";
		strObjName += $(objSender).parent().parent().find("span[id$='Label4']").text();*/
		
	if(!confirm("確定要刪除" + strObjName + "？")){
		return;
	}
	doDelUserAuth(objSender);
}

function doDelUserAuth(objSender){
	var strUserID = $(objSender).parent().parent().parent().parent().parent().parent().parent().prev().find(":hidden:eq(0)").val();
    //var strRecID = $(objSender).parent().parent().find(":hidden[id$='hidRecID']").val(); 
    var strRecID = $(objSender).parent().parent().find(":hidden:eq(0)").val(); // feng 20170321 
	$("#divLoading").remove();
	$("#powerlist").append("<div style='position:absolute;top:0px;left:0px;background-color:#000;filter:alpha(opacity=45);-moz-opacity: 0.75;opacity: 0.75;color:gray;width:100%;height:100%;border-width:1px;border-style:solid;border-color:white;text-align:center;font-size=18px' id='divLoading'><br><br><br><br><br><br><br><br>Deleting data............<br><br><img src='../../../App_Themes/pwfBody/images/AjaxLoadingBar1.gif'></div>");
	$.ajax(
		{
			url:"../../../Pub/RemoteMethods/WorkFlowSequence.aspx",
			type:"post",
			data:"EventName=delAuth&UserID="+strUserID+"&RecID="+strRecID,
			dataType:"text",
			timeout:100000,
			error:function(err){
				objServerVar = {
					IsOK:false,
					ServerMsg : ""
				}
				objServerVar.ServerMsg = "遠端連線發生錯誤或連線逾時，請再試一次，如果持續發生請連絡相關人員！！";
			},
			success:function(strServerVar){
				eval("objServerVar ="+strServerVar);
			},
			complete:function(){
				alert(objServerVar.ServerMsg);
				$("#divLoading").remove();
				if(objServerVar.IsOK){
					$(objSender).parent().parent().parent().parent().parent().parent().html( "<br>" + objServerVar.Result + "<br>");
				}
			}
			
		}
	);	
	
}

