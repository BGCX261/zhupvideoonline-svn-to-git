<!--#include file="../include/head.asp"-->
<%
Const EnableUploadFile="Yes"        '�Ƿ񿪷��ļ��ϴ�
Const MaxFileSize=2000        '�ϴ��ļ���С����
Const SaveUpFilesTopPic="../images"        '����ϴ��ļ���Ŀ¼
Const UpFileTypeTopPic="jpg"        '������ϴ��ļ�����
%>

<!--#include file="upload.asp"-->
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
SavePath = SaveUpFilesTopPic   '����ϴ��ļ���Ŀ¼
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
	If User_Group<>"admin" then
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
		Forumupload=split(UpFileTypeTopPic,"|")
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
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileTypeTopPic
			founderr=true
		end if
		
		dim strJS
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			filename=SavePath&file.FileName'&"."&fileExt

			file.SaveToFile Server.mappath(FileName)   '�����ļ�

			msg="�ϴ��ļ��ɹ���"
			

			
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "history.go(-1);" & vbcrlf		
	  	strJS=strJS & "</script>"
	  	response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>
