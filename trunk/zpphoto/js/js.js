<!--缩略图脚本
var flag=false;
function DrawImage(ImgD){
var image=new Image();
image.src=ImgD.src;
if(image.width>0 && image.height>0){
flag=true;
if(image.width/image.height>= 100/120){
if(image.width>100){ 
ImgD.width=100;
ImgD.height=(image.height*100)/image.width;
}else{
ImgD.width=image.width;
ImgD.height=image.height;
}
ImgD.alt=" 点击查看 ";
}
else{
if(image.height>120){ 
ImgD.height=120;
ImgD.width=(image.width*120)/image.height; 
}else{
ImgD.width=image.width; 
ImgD.height=image.height;
}
ImgD.alt=" 点击查看 ";
}
}
} 
//-->

<!--弹出窗口脚本
function winopen(url)                    
     {                    
        window.open(url,"search","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=yes,top=35 ,left=100");          
      }
//-->

<!--输入检查
function checkinput()
{
	if (document.search.keyword.value=="")
	 {	
		alert("请输入想查询的内容！");
		document.search.keyword.focus();
		
		return false;
	 }
	 return true;
}
-->

