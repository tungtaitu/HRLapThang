var obj;

// Creates the XMLHTTP Request object
function getXMLHTTPRequest()
{
	var xRequest=null;
	if (window.XMLHttpRequest)
	{
		xRequest = new XMLHttpRequest();
	}
	else if (typeof ActiveXObject != "undefined")
	{
		try 
		{      
            xRequest = new ActiveXObject("Msxml2.XMLHTTP");
        }
        catch (exp1) 
        {
            try 
            {        
                xRequest = new ActiveXObject("Microsoft.XMLHTTP");
            }
            catch (exp2) 
            {
                xRequest = false;
            }
        }
    }
	
	return xRequest;
}

$(document).ready(function() {
						   
	var hash = window.location.hash.substr(1);
	var href = $('#nav li a').each(function(){
		var href = $(this).attr('href');
		if(hash==href.substr(0,href.length-5)){
			var toLoad = hash+'.html #content';
			$('#content').load(toLoad)
		}											
	});

	$('#btnQuery').click(function(){
		obj=getXMLHTTPRequest();
						  
		var toLoad = $(this).attr('href')+' #content';
		$('#content').hide('fast',loadContent);
		$('#load').remove();
		$('#wrapper').append('<span id="load">Loading data, please wait ....</span> ');
		$('#load').fadeIn('normal');
		//window.location.hash = $(this).attr('href').substr(0,$(this).attr('href').length-5);
		function loadContent() {
			if(obj != null)
			{
				if (obj.readyState == 4)
				{
					if (obj.status == 200)
					{
						$('#content').load(toLoad,'',showNewContent())
					}
				}
			}
			
		}
		function showNewContent() {
			$('#content').show('normal',hideLoader());
		}
		function hideLoader() {
			$('#load').fadeOut('normal');
		}
		return false;
		
	});

});