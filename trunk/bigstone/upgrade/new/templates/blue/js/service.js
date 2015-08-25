document.getElementById("service").style.top = 150 + "px";
window.onscroll=function()
{
	var scrollTop = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0;
	document.getElementById("service").style.top = scrollTop + 150 + "px";
}
function show_service()
{
	if(document.getElementById("service").style.width == "33px")
	{
		document.getElementById("service").style.width = "143px";
	}else{
		document.getElementById("service").style.width = "33px";
	}
}
//新秀