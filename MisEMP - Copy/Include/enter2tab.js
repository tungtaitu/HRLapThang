document.onkeydown=function(evt)
 {
   var isie = (document.all) ? true : false; 
   var key; 
   var srcobj;
   if (isie) 
   {
     key = event.keyCode; 
     srcobj=event.srcElement;
     }
   else 
   {
     key = evt.which; 
     srcobj=evt.target;
     }               
     if(key==13 && srcobj.type!='button' && srcobj.type!='submit' &&srcobj.type!='reset' 
       && srcobj.type!='textarea' && srcobj.type!='' )                 
     {
       if(isie)
       {
        event.keyCode=9;
       }
       else
       {                        
         var el=getNextElement(evt.target);
                 if (el.type!='hidden')
                    el.focus(); 
                 else
                    while (el.type=='hidden')
                       el=getNextElement(el); 
                    el.focus(); 
                 return false;
       }
     }
  }  
  
function getNextElement (field) {
   var form = field.form;
   for (var e = 0; e < form.elements.length; e++) { 
	 if (field == form.elements[e])
		 break;
   }
   return form.elements[++e % form.elements.length];
 }   