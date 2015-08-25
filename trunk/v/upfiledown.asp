<!--#include file="Inc/config.asp"-->

<!--#include file="Inc/upload.asp"-->
<%
const upload_type=0   '上传方法：0=无惧无组件上传类，1=FSO上传 2=lyfupload，3=aspupload，4=chinaaspupload

dim upload,file,formName,SavePath,filename,fileExt
dim upNum
dim EnableUpload
dim Forumupload
dim ranNum
dim uploadfiletype
dim msg,founderr
msg=""
founderr=false
EnableUpload=false
SavePath = SaveUpFilesPathDown   '存放上传文件的目录
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<%
if EnableUploadFile="No" then
	response.write "系统未开放文件上传功能"
else
	if session("adminName")="" then
		response.Write("请登录后再使用本功能！")
	else
		select case upload_type
			case 0
				call upload_0()  '使用化境无组件上传类
			case else
				'response.write "本系统未开放插件功能"
				'response.end
		end select
	end if
end if
%>
</body>
</html>
<%
sub upload_0()    '使用化境无组件上传类
	set upload=new upload_file    '建立上传对象
	for each formName in upload.file '列出所有上传了的文件
		set file=upload.file(formName)  '生成一个文件对象
		if file.filesize<100 then
 			msg="请先选择你要上传的文件！"
			founderr=true
		end if
		if file.filesize>(MaxFileSizeDown*1024) then
 			msg="文件大小超过了限制，最大只能上传" & CStr(MaxFileSizeDown) & "K的文件！"
			founderr=true
		end if

		fileExt=lcase(file.FileExt)
		Forumupload=split(UpFileTypeDown,"|")
		for i=0 to ubound(Forumupload)
			if fileEXT=trim(Forumupload(i)) then
				EnableUpload=true
				exit for
			end if
		next
		if fileEXT="asp" or fileEXT="asa" or fileEXT="aspx" then
			EnableUpload=false
		end if
		if EnableUpload=false then
			msg="这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileTypeDown
			founderr=true
		end if
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

			file.SaveToFile Server.mappath(FileName)   '保存文件

			msg="上传文件成功！"
			
			strJS=strJS & "parent.Uploaddown.focus();" & vbcrlf 	  		
			strJS=strJS & "var range = parent.Uploaddown.value;" & vbcrlf
		  	strJS=strJS & "range='[upload=" & FileType & "]" & FileName & "[/upload]';" & vbcrlf
			'strJS=strJS & "parent.myform.IncludePic.checked=true;" & vbcrlf
	  		'strJS=strJS & "parent.myform.DefaultPicUrl.value='" & FileName & "';" & vbcrlf
			'strJS=strJS & "parent.myform.DefaultPicList.options[parent.myform.DefaultPicList.length] = new Option('" & filename & "','" & filename & "');" & vbcrlf
			'strJS=strJS & "parent.myform.DefaultPicList.selectedIndex+=1;" & vbcrlf
			strJS=strJS & "if(parent.myform.Uploaddown.value==''){" & vbcrlf
			strJS=strJS & "parent.myform.Uploaddown.value+='" & filename & "';}" & vbcrlf
			strJS=strJS & "else{" & vbcrlf & "parent.myform.Uploaddown.value='" & filename & "';}" & vbcrlf
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "history.go(-1);" & vbcrlf
		strJS=strJS & "parent.Uploaddown.focus();" & vbcrlf 	  		
	  	strJS=strJS & "</script>"
	  	response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>
