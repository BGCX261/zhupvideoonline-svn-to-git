<!--#include file="Inc/config.asp"-->

<!--#include file="Inc/upload.asp"-->
<%
const upload_type=0   '�ϴ�������0=�޾�������ϴ��࣬1=FSO�ϴ� 2=lyfupload��3=aspupload��4=chinaaspupload

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
SavePath = SaveUpFilesPath   '����ϴ��ļ���Ŀ¼
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<%
if EnableUploadFile="No" then
	response.write "ϵͳδ�����ļ��ϴ�����"
else
	if session("adminName")="" then
		response.Write("���¼����ʹ�ñ����ܣ�")
	else
		select case upload_type
			case 0
				call upload_0()  'ʹ�û���������ϴ���
			case else
				'response.write "��ϵͳδ���Ų������"
				'response.end
		end select
	end if
end if
%>
</body>
</html>
<%
sub upload_0()    'ʹ�û���������ϴ���
	set upload=new upload_file    '�����ϴ�����
	for each formName in upload.file '�г������ϴ��˵��ļ�
		set file=upload.file(formName)  '����һ���ļ�����
		if file.filesize<100 then
 			msg="����ѡ����Ҫ�ϴ����ļ���"
			founderr=true
		end if
		if file.filesize>(MaxFileSize*1024) then
 			msg="�ļ���С���������ƣ����ֻ���ϴ�" & CStr(MaxFileSize) & "K���ļ���"
			founderr=true
		end if

		fileExt=lcase(file.FileExt)
		Forumupload=split(UpFileType,"|")
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
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileType
			founderr=true
		end if
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

			file.SaveToFile Server.mappath(FileName)   '�����ļ�

			msg="�ϴ��ļ��ɹ���"
			
			strJS=strJS & "parent.HtmlEdit.focus();" & vbcrlf 	  		
			strJS=strJS & "var range = parent.HtmlEdit.document.selection.createRange();" & vbcrlf
			FileType=right(fileExt,3)
			select case FileType
			    case "jpg","gif","png","bmp"
					strJS=strJS & "range.pasteHTML('<img src=" & filename & ">');" & vbcrlf
				case "swf"
					strJS=strJS & "range.pasteHTML('<object classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0>"
		  			strJS=strJS & "<param name=movie value=" & FileName & ">"
		  			strJS=strJS & "<param name=quality value=high>"
		  			strJS=strJS & "<embed src=" & FileName & " quality=high pluginspage=http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash type=application/x-shockwave-flash>"
					strJS=strJS & "</embed></object>');" & vbcrlf
				case else
		  			strJS=strJS & "range.text='[upload=" & FileType & "]" & FileName & "[/upload]';" & vbcrlf
	  		end select
			strJS=strJS & "parent.parent.myform.IncludePic.checked=true;" & vbcrlf
	  		strJS=strJS & "parent.parent.myform.DefaultPicUrl.value='" & FileName & "';" & vbcrlf
			strJS=strJS & "parent.parent.myform.DefaultPicList.options[parent.parent.myform.DefaultPicList.length] = new Option('" & filename & "','" & filename & "');" & vbcrlf
			strJS=strJS & "parent.parent.myform.DefaultPicList.selectedIndex+=1;" & vbcrlf
			strJS=strJS & "if(parent.parent.myform.UploadFiles.value==''){" & vbcrlf
			strJS=strJS & "parent.parent.myform.UploadFiles.value+='" & filename & "';}" & vbcrlf
			strJS=strJS & "else{" & vbcrlf & "parent.parent.myform.UploadFiles.value+='|'+'" & filename & "';}" & vbcrlf
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "history.go(-1);" & vbcrlf
		strJS=strJS & "parent.HtmlEdit.focus();" & vbcrlf 	  		
	  	strJS=strJS & "</script>"
	  	response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>