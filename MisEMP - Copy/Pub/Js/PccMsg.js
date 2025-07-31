function PccMsg()
{
    var text ="<PccMsg></PccMsg>";
    if(arguments.length == 1)
    {
         var text = arguments[0];
         var tab = text.substring(0,8);
         if (tab != "<PccMsg>")// Load file XMl
         {
             if (window.XMLHttpRequest)// Firefox
                {
                    xhttp = new XMLHttpRequest()
                }
                else //Internet Explorer
                {
                    xhttp = new ActiveXObject("Microsoft.XMLHTTP")
                }
                xhttp.open("GET",text,false);
                xhttp.send("");
                this.xmlDoc = xhttp.responseXML;
                return;
         }
    } 
        
    if (window.DOMParser)// Firefox
    {
        parser = new DOMParser();
        this.xmlDoc = parser.parseFromString(text,"text/xml");
    }
    else // Internet Explorer
    {
      this.xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
      this.xmlDoc.async="false";
      this.xmlDoc.loadXML(text);
    }
}

PccMsg.prototype.GetXmlStr=function(){
    if (window.DOMParser)// Firefox
    {
       return (new XMLSerializer()).serializeToString(this.xmlDoc);//Convert XML Document to String
    }
    else // Internet Explorer
    {
        return this.xmlDoc.xml;
    } 
}
PccMsg.prototype.CreateFirstNode = function(strTag, strValue){
    if (window.DOMParser)// Firefox
    {
        var node = this.xmlDoc.createElement(strTag);
        node.textContent = strValue;
        this.xmlDoc.getElementsByTagName("PccMsg")[0].appendChild(node);
    }
    else // Internet Explorer
    {
        var node = this.xmlDoc.createElement(strTag);
        node.text = strValue;
        this.xmlDoc.getElementsByTagName("PccMsg")[0].appendChild(node);
    } 
}
PccMsg.prototype.Query = function (xmlPath) {
    var value = "";
    var path = "PccMsg/" + xmlPath;

    if (window.DOMParser)// Firefox
    {

        var node = SelectSingleNode(this.xmlDoc, path);
        if (node != null) {
            value = node.textContent;
        }
    }
    else {

        var node = this.xmlDoc.selectSingleNode(path);
        if (node != null) {
            value = node.text;
        }

    }

    return value;
}

PccMsg.prototype.QueryNode = function(xmlPath) {
    var value = "";
    var path = "PccMsg/" + xmlPath;
	
    if (window.DOMParser)// Firefox
    {
		var xpe = new XPathEvaluator();
        var nsResolver = xpe.createNSResolver(this.xmlDoc.ownerDocument == null ? this.xmlDoc.documentElement : this.xmlDoc.ownerDocument.documentElement);
        var results = xpe.evaluate(path, this.xmlDoc, nsResolver, 0, null); 
        
        
        var found = [];
		var res;
		while (res = results.iterateNext())
			found.push(res);
		return found;
        
        return evaluateXPath(this.xmlDoc,path);
    }
    else {

        return this.xmlDoc.selectNodes(path);
    }

    return value;
}

function SelectSingleNode(xmlDoc, elementPath) {
    if (window.ActiveXObject) {
        return xmlDoc.selectSingleNode(elementPath);
    }
    else {
        var xpe = new XPathEvaluator();
        var nsResolver = xpe.createNSResolver(xmlDoc.ownerDocument == null ? xmlDoc.documentElement : xmlDoc.ownerDocument.documentElement);
        var results = xpe.evaluate(elementPath, xmlDoc, nsResolver, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
        return results.singleNodeValue;
    }
}
/*
function evaluateXPath(aNode, aExpr) {
	  var xpe = new XPathEvaluator();
	  var nsResolver = xpe.createNSResolver(aNode.ownerDocument == null ?
	    aNode.documentElement : aNode.ownerDocument.documentElement);
	  var result = xpe.evaluate(aExpr, aNode, nsResolver, 0, null);
	  var found = [];
	  var res;
	  while (res = result.iterateNext())
	    found.push(res);
	  return found;
	}

*/
