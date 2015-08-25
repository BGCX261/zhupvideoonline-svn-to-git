<%@language=vbscript enablesessionstate=false lcid=1033%>

<!-- #include file="Progress_UploadJ.asp" -->

<%
	Response.CacheControl = "no-cache"
	Response.Expires = -1

	progressID = Request.QueryString("ProgressID")

	Set upload = New DoteyUpload
	Set Info = upload.GetProgressInfo(progressID)

If Info.ReadyState = "loaded" Then
%>
文件上传结束
<script>window.close();</script>
<%
Else
%>
<html>
<head>
	<meta http-equiv="refresh" content="1;url=Progress_WinJ.asp?ProgressID=<%=ProgressID%>">
	<title>已上传<%= Info.Percentage %></title>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0">
	<tr><td>总 大 小:&nbsp;</td><td><%= Info.TotalSize %></td></tr>
	<tr><td>已经上传:&nbsp;</td><td><%= Info.SizeCompleted %></td></tr>
	<tr><td>平均速率:</td><td><%= Info.TransferRate %></td></tr>
	<tr><td>使用时间:</td><td><%= Info.ElapsedTime %></td></tr>
	<tr><td>剩余时间:</td><td><%= Info.TimeLeft %></td></tr>
</table>
<table border="3" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr><td width="<%= Info.Percentage %>" bgcolor="blue"></td><td>&nbsp;</td></tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
End If
%>