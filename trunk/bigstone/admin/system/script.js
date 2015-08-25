var sinsiu,njb,interval;

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

//////////////////////////////////////////////////////////////
function show_box(id,w,h)
{
	var st = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0;
	cw = document.body.clientWidth;
	t = st + 150;
	l = Math.floor((cw - w - 10) / 2);
	document.getElementById(id).style.display = "block";
	document.getElementById(id).style.width = w + "px";
	document.getElementById(id).style.height = h + "px";
	document.getElementById(id).style.top = t + "px";
	document.getElementById(id).style.left = l + "px";
}
function hid_box(id)
{
	document.getElementById(id).style.display = "none";
}
function result()
{
	show_box("result",250,95);
}
function get_result(fil)
{
	ajax("post","../system/ajax.asp","cmd=result",
	function(data)
	{
		if(data == 1)
		{
			result();
			clearInterval(interval);
		}
	});
}
function del(id)
{
	document.getElementById("del_id").value = id;
	var st = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0;
	cw = document.body.clientWidth;
	t = st + 150;
	l = Math.floor((cw - 260) / 2);
	document.getElementById("del").style.display = "block";
	document.getElementById("del").style.width = "300px";
	document.getElementById("del").style.height = "125px";
	document.getElementById("del").style.top = t + "px";
	document.getElementById("del").style.left = l + "px";
}

function get_value(name)
{
	return document.getElementsByName(name).item(0).value;
}
function set_value(name,val)
{
	document.getElementsByName(name).item(0).value = val;
}

//////////////////////////////////////////////////////////////
function do_upload(val)
{
	var dir,fil;
	if(val == "upl_img")
	{
		dir = get_value("upl_img_dir");
		fil = get_value("upl_img_path");
	}else{
		dir = get_value("upl_fil_dir");
		fil = get_value("upl_fil_path");
	}
	if(fil != "")
	{
		var site = fil.lastIndexOf(".");
		if(site != -1)
		{
			site = fil.lastIndexOf("\\");
			fil = fil.substr(site + 1,fil.length - site);
			ajax("post","../system/ajax.asp","cmd=get_name&dir="+dir+"&fil="+fil,
			function(data)
			{
				fil = data;
				if(val == "upl_img")
				{
					document.upl_img_form.action = "upload.asp?dir=" + dir + "&fil=" + fil + "&reduce=1";
					document.upl_img_form.submit();
					get_pic();
				}else{
					document.upl_fil_form.action = "upload.asp?dir=" + dir + "&fil=" + fil + "&reduce=0";
					document.upl_fil_form.submit();
					get_fil();
				}
				hid_box(val);
			});
		}
	}
}
function get_pic()
{
	njb = setInterval("get_pic_ajax()",1000);
}
function get_pic_ajax()
{
	ajax("post","../system/ajax.asp","cmd=get_pic",
	function(data)
	{
		if(data != "")
		{
			document.getElementById("get_pic").innerHTML = "<img src='" + data + "' height='100px' />";
			set_value("pic_path",data);
		}
	});
}
function set_seo(val)
{
	var key = get_value(val+"_keywords");
	var des = get_value(val+"_description");
	set_value("seo_key",key);
	set_value("seo_des",des);
	show_box('seo',430,220);
}
function get_seo(val)
{
	var key = get_value("seo_key");
	var des = get_value("seo_des");
	set_value(val+"_keywords",key);
	set_value(val+"_description",des);
	hid_box("seo");
}
function get_fil()
{
	njb = setInterval("get_fil_ajax()",1000);
}
function get_fil_ajax()
{
	ajax("post","../system/ajax.asp","cmd=get_fil",
	function(data)
	{
		if(data != "")
		{
			document.getElementById("get_fil").innerHTML = data;
			set_value("fil_path",data);
		}
	});
}
function set_web_url()
{
	var fil_url = get_value("fil_url");
	ajax("post","../system/ajax.asp","cmd=set_fil&fil_url="+fil_url,
	function(data)
	{
		if(data == 1)
		{
			document.getElementById("get_fil").innerHTML = fil_url;
			set_value("fil_path",fil_url);
			hid_box('web_fil');
		}
	});
}

function show_edit(val)
{
	document.getElementById(val+"_1").style.display = "none";
	document.getElementById(val+"_2").style.display = "block";
	document.getElementById(val+"_2").childNodes.item(0).focus();
}
function hid_edit(val)
{
	document.getElementById(val+"_1").style.display = "block";
	document.getElementById(val+"_2").style.display = "none";
}
function set_index(id,val,cat)
{
	ajax("post","deal.asp","cmd=set_"+cat+"_index&id="+id+"&index="+val,
	function(data)
	{
		if(data == 1)
		{
			document.getElementById("index_"+id+"_1").innerHTML = document.getElementById("index_"+id).value + "<img src='../system/images/pencil.gif'>";
			hid_edit("index_"+id);
		}
	});
}
function set_top(id,val,cat)
{
	if(val) val = 1;else val = 0;
	ajax("post","deal.asp","cmd=set_"+cat+"_top&id="+id+"&top="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function set_show(id,val,cat)
{
	if(val) val = 1;else val = 0;
	ajax("post","deal.asp","cmd=set_"+cat+"_show&id="+id+"&show="+val,
	function(data)
	{
		if(data == 1)result();
		if(data == 2)show_box('no_deal',300,95);
	});
}

//新秀