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
SavePath = SaveUpFilesPathDown   '����ϴ��ļ���Ŀ¼
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
		if file.filesize>(MaxFileSizeDown*1024) then
 			msg="�ļ���С���������ƣ����ֻ���ϴ�" & CStr(MaxFileSizeDown) & "K���ļ���"
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
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileTypeDown
			founderr=true
		end if
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

			file.SaveToFile Server.mappath(FileName)   '�����ļ�

			msg="�ϴ��ļ��ɹ���"
			
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
