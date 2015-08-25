<style type="text/css">
<!--
BODY{
font-size:12px;
}
-->
</style>
<!-- #include file="Progress_Upload.asp" -->

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

  'Response.Write("您已上传 " & Upload.Files.Count & " 个文件到: " & path & "<hr>")
  for each File in Upload.Files.Items
    Response.Write("文件名: " & File.FileName & "<br>")
    Response.Write("文件大小: " & FormatNumber(File.FileSize/1024,2) & " KB<br>")
    'Response.Write("文件类型: " & File.FileType & "<br>")
    Response.Write("客户端路径: " & File.FilePath & "<br>")
	strJS="<SCRIPT language=javascript>" & vbcrlf			
			'strJS=strJS & "parent.Uploaddown.focus();" & vbcrlf 	  		
			'strJS=strJS & "var range = parent.Uploaddown.value;" & vbcrlf
			strJS=strJS & "parent.myform.downSize.value='" & FormatNumber(File.FileSize/1024,2) & "';" & vbcrlf
		  	'strJS=strJS & "range='" & FileName & "';" & vbcrlf
			strJS=strJS & "if(parent.myform.Uploaddown.value==''){" & vbcrlf
			strJS=strJS & "parent.myform.Uploaddown.value+='" & path & "/" & File.FileName & "';}" & vbcrlf
			strJS=strJS & "else{" & vbcrlf & "parent.myform.Uploaddown.value='" & path & "/" & File.FileName & "';}" & vbcrlf
		'strJS=strJS & "alert('上传成功?! ');" & vbcrlf
	  	'strJS=strJS & "history.go(-1);" & vbcrlf
		strJS=strJS & "parent.Uploaddown.focus();" & vbcrlf 	  		
	  	strJS=strJS & "</script>"
	  	response.write strJS
  next
  
  		
%>