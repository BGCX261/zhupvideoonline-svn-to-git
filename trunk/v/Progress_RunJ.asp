<style type="text/css">
<!--
BODY{
font-size:12px;
}
-->
</style>
<!-- #include file="Progress_UploadJ.asp" -->

<%
  Server.ScriptTimeout = 3600
 
  path = "UploadFilesDown" 

  set Upload = new DoteyUpload

  Upload.ProgressID = Request.QueryString("ProgressID")
  Upload.SaveTo(path)

  if Upload.ErrMsg <> "" then 
    Response.Write(Upload.ErrMsg)
    Response.End()
  end if

  if Upload.Files.Count > 0 then
    Items = Upload.Files.Items
  end if

  'Response.Write("�����ϴ� " & Upload.Files.Count & " ���ļ���: " & path & "<hr>")
  datestr=Now()
  datestr=replace(datestr," ", "")'ȥ���ո� 
  datestr=replace(datestr,"/", "")'ȥ��/ 
  datestr=replace(datestr,":", "")'ȥ��:
  for each File in Upload.Files.Items
    Response.Write("�ļ���: " & datestr & "_" & File.FileName & "<br>")
    Response.Write("�ļ���С: " & FormatNumber(File.FileSize/1024,2) & " KB<br>")
    'Response.Write("�ļ�����: " & File.FileType & "<br>")
    Response.Write("�ͻ���·��: " & File.FilePath & "<br>")
	strJS="<SCRIPT language=javascript>" & vbcrlf			
			'strJS=strJS & "parent.Uploaddown.focus();" & vbcrlf 	  		
			'strJS=strJS & "var range = parent.Uploaddown.value;" & vbcrlf
			strJS=strJS & "parent.myform.downSizeJ.value='" & FormatNumber(File.FileSize/1024,2) & "';" & vbcrlf
		  	'strJS=strJS & "range='" & FileName & "';" & vbcrlf
			strJS=strJS & "if(parent.myform.UploaddownJ.value==''){" & vbcrlf
			strJS=strJS & "parent.myform.UploaddownJ.value+='" & path & "/"& datestr & "_" & File.FileName & "';}" & vbcrlf
			strJS=strJS & "else{" & vbcrlf & "parent.myform.UploaddownJ.value='" & path & "/" & datestr & "_" & File.FileName & "';}" & vbcrlf
		'strJS=strJS & "alert('�ϴ��ɹ�?! ');" & vbcrlf
	  	'strJS=strJS & "history.go(-1);" & vbcrlf
		strJS=strJS & "parent.UploaddownJ.focus();" & vbcrlf 	  		
	  	strJS=strJS & "</script>"
	  	response.write strJS
  next
  
  		
%>