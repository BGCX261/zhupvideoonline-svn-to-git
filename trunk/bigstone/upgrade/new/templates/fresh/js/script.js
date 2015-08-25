function ajax(type,file,text,func)
{
	var XMLHttp_Object;
	try{XMLHttp_Object = new ActiveXObject("Msxml2.XMLHTTP");}
	catch(new_ieerror)
	{
		try{XMLHttp_Object = new ActiveXObject("Microsoft.XMLHTTP");}
		catch(ieerror){XMLHttp_Object = false;}
	}
	if(!XMLHttp_Object && typeof XMLHttp_Object != "undefiend")
	{
		try{XMLHttp_Object = new XMLHttpRequest();}
		catch(new_ieerror){XMLHttp_Object = false;}
	}
	type = type.toUpperCase();
	if(type == "GET") file = file + "?" + text;
	XMLHttp_Object.open(type,file,true);
	if(type == "POST") XMLHttp_Object.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	XMLHttp_Object.onreadystatechange = function ResponseReq()
	{
		if(XMLHttp_Object.readyState == 4) func(XMLHttp_Object.responseText);
	};
	if(type == "GET") text = null;
	XMLHttp_Object.send(text);
}

function picresize(obj,MaxWidth,MaxHeight)
{
	if(!window.XMLHttpRequest)
	{
		obj.onload = null;
		img = new Image();
		img.src = obj.src;
		if (img.width > MaxWidth && img.height > MaxHeight)
		{
			if(img.width / img.height > MaxWidth / MaxHeight)
			{
				obj.height = MaxWidth * img.height / img.width;
				obj.width = MaxWidth;
			}else{
				obj.width = MaxHeight * img.width / img.height;
				obj.height = MaxHeight;
			}
		}else if(img.width > MaxWidth){
			obj.height = MaxWidth * img.height / img.width;
			obj.width = MaxWidth;
		}else if(img.height > MaxHeight){
			obj.width = MaxHeight * img.width / img.height;
			obj.height = MaxHeight;
		}else{
			obj.width = img.width;
			obj.height = img.height;
		}
	}
}

//新秀