<html>
<head>
<%
if request.cookies("name49s")="" then
  response.redirect "../login.asp"
end if
%>
<%
pathset = 0
upset = 2
nameset = 2
response.cookies("pic1")=""
response.cookies("pic2")=""
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>作品上传</title>
<style type="text/css">
body { color:#000; font-size:12px; text-align:center;}
.cont{width:700px; margin:0 auto;text-align:left;}
#FileList{ border:1px #999 solid; padding:6px; background:#efefef; font-size:12px; font-family:宋体}
#FileList form{ margin:0; padding:0}
.list{ border:1px #ddd solid; background:#fff; padding:3px; margin:2px 0; line-height:18px}
.tips{ color:red}
.almText{ border:#999 solid 1px; background:#FFC; line-height:180%; padding:6px;}
.aCenter{ text-align:center; padding:6px;}
.bText{ font-weight:bold; font-size:14px;}
</style>
<script type="text/javascript">
/*
Created by anlige，1034555083，zhanghuiguoanlige@126.com
Modify by zhup
*/
var fileCount=0;
var fileQuene=[];
var currentIndex=0;
function doSomething(fil){
	if(fil!=null){
		fileCount ++;
		if(fileCount>=21) return alert("一次最多只允许上传20个文件！");
		var list = fil.parentNode.parentNode;
		list.id="file_" + fileCount;
		
		var frm = fil.parentNode;
		frm.className = "hasFile";
		frm["pid"].value=fileCount;
		frm.style.display="none";
		
		var newX = document.createElement("div");
		newX.innerHTML="<input type=\"hidden\" id=\"tabs" + fileCount + "\" name=\"tabs" + fileCount + "\" >";
		document.getElementById("frm").appendChild(newX);
		
		var newV = document.createElement("div");
		var fname=fil.value.replace(/\//igm,"\\").split("\\");
		newV.innerHTML="<div class=\"filename\">" + fname[fname.length-1] + "</div><div class=\"tips\" id=\"tabp" + fileCount + "\">等待上传...</div>";
		list.appendChild(newV);
		
	}
	var newFile = document.createElement("div");
	newFile.className = "list";
	newFile.innerHTML="<form action=\"upload.asp\" method=\"post\" enctype=\"multipart/form-data\" target=\"uploadfrm\"><input onchange=\"doSomething(this);\" class=\"file\" type=\"file\" name=\"file\" /><input type=\"hidden\" name=\"pid\" value=\"0\" /> <input type=\"hidden\" name=\"que\" value=\"0\" /> <input class=\"but\" type=\"button\" onclick=\"doUpload();\" value=\"上传\" /></form>";
	document.getElementById("FileList").appendChild(newFile);
}
function doUpload(){
	var files = document.getElementById("FileList");
	var forms = files.getElementsByTagName("form");
	if(forms.length==1){alert("请添加文件");return;}
	var k=0;
	for(var i=0;i<forms.length;i++){
		var frm = forms[i];
		if(frm.className=="hasFile"){
			frm["que"].value=k;
			k++;
			fileQuene.push(frm);
		}
	}
	_doUpload(); 
}
function _doUpload(){
	if(currentIndex>=fileQuene.length){
		document.getElementById("FileList").lastChild.innerHTML="全部上传完毕";
		fileQuene=[];return;
	}
	document.getElementById("tabp"+fileQuene[currentIndex]["pid"].value).value="正在上传...";
	fileQuene[currentIndex].submit();
	
}
function setStep(iserr,desc){
	var obj = document.getElementById("tabs"+fileQuene[currentIndex]["pid"].value);
	var objP = document.getElementById("tabp"+fileQuene[currentIndex]["pid"].value);
	
	if(iserr){
		desc = "上传失败：" + desc;
	}else{
		objP.value= "上传成功：" + desc ;
		objP.style.color="green";
		objP.style.fontWeight="bold";
	}
	objP.innerHTML=desc;
	obj.value="upload/imgSwf/" + desc;
	currentIndex++;
	_doUpload(currentIndex);
	
}
function okUp(){
	n.submit();
	}
window.onload=function(){doSomething(null);}
</script>
</head>
<body>
<div class="cont">
<div class="almText">
<ul>
    <li><span class="bText">图片上传</span></li>
    <li>若要上传本地图片，请先上传图片，点击“上传”可查看完成后的状态。再点击“下一步”。图片地址会自动对应到图片地址栏</li>
    <li>请上传符合要求的文件，可同时上传多个文件（最多20个）</li>
    <li>若要添加网络外链图片，请 <a href="../addPhoto.asp">点这里</a> 直接添加，在对应的图片地址框里填入外链图片地址即可。</li>
</ul>
</div>
<form method="POST" action="../addPhoto.asp" id="n"><div id="frm"></div>

</form>
<div id="FileList"></div>
<iframe src="about:blank" width="0" height="0" frameborder="0" name="uploadfrm"></iframe>
<div class="aCenter"><input type="button" value="保存上传图片" onClick="okUp()"> <input type="button" value="返 回" onClick="window.location='../adminMain.asp'"> <input type="button" value="添加外链图片" onClick="window.location='../addPhoto.asp'"></div>
</div>
</body>
</html>

