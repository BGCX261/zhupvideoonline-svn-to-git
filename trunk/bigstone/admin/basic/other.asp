<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%dim checked_1,checked_0%>
<div class="here">首页公司简介设置</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%">是否过滤HTML标签</td>
    <td>
    <%dim about_filter:about_filter = int(get_config("about_filter"))%>
    <input type="radio" name="filter" onClick="set_about_filter(1)" <%if about_filter = 1 then echo "checked"%> />是
    <input type="radio" name="filter" onClick="set_about_filter(0)" <%if about_filter = 0 then echo "checked"%> />否
    </td>
  </tr>
  <tr>
    <td>截取字符串长度：</td>
    <td><input type="text" style="width:50px;" onBlur="set_about_len(this.value)" value="<%=get_config("about_len")%>" /> 个字符</td>
  </tr>
</table>
</div>
<div class="here">列表长度设置</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%">首页文章列表长度：</td>
    <td><input type="text" style="width:50px;" onBlur="set_list_len('index_list_len',this.value)" value="<%=get_config("index_list_len")%>" /> 行</td>
  </tr>
  <tr>
    <td>图片页列表长度：</td>
    <td><input type="text" style="width:50px;" onBlur="set_list_len('img_list_len',this.value)" value="<%=get_config("img_list_len")%>" /> 张</td>
  </tr>
  <tr>
    <td>文章页列表长度：</td>
    <td><input type="text" style="width:50px;" onBlur="set_list_len('art_list_len',this.value)" value="<%=get_config("art_list_len")%>" /> 行</td>
  </tr>
</table>
</div>
<div class="here">图片设置</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%">AspJpeg组件缩略图：</td>
    <td>
    最大宽度：<input type="text" style="width:50px;" onBlur="set_reduce_img('width',this.value)" value="<%=get_config("reduce_img_width")%>" /> px &nbsp;&nbsp;
    最大高度：<input type="text" style="width:50px;" onBlur="set_reduce_img('height',this.value)" value="<%=get_config("reduce_img_height")%>" /> px
    </td>
  </tr>
</table>
</div>
<div class="here">会话方式设置</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%">选择会话方式：</td>
    <td>
    <input type="radio" name="session" onClick="set_session_method(1)" <%if S_SESSION = 1 then echo "checked"%> /> SESSION
    <input type="radio" name="session" onClick="set_session_method(0)" <%if S_SESSION = 0 then echo "checked"%> /> COOKIES
    </td>
  </tr>
</table>
</div>
<div class="here">模板编译模式设置</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%">选择模板编译模式：</td>
    <td>
    <%
	dim sinty_xml:sinty_xml = server.mappath(S_TPL_PATH&"sinty.xml")
	dim xml:set xml = server.createobject("microsoft.xmldom")
	xml.load(sinty_xml)
	compile = int(xml.selectsinglenode("site").getattribute("compile"))
	%>
    <input type="radio" name="compile" onClick="set_compile(0)" <%if compile = 0 then echo "checked"%> /> 不编译
    <input type="radio" name="compile" onClick="set_compile(1)" <%if compile = 1 then echo "checked"%> /> 只编译未编译的模板
    <input type="radio" name="compile" onClick="set_compile(2)" <%if compile = 2 then echo "checked"%> /> 总是编译
    </td>
  </tr>
</table>
</div>
<div class="here">静态设置</div>
<div class="block">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%">静态模式：</td>
    <td>
    <input type="radio" name="static" onClick="set_static(0)" <%if S_STATIC = 0 then echo "checked"%> />动态
    <input type="radio" name="static" onClick="set_static(1)" <%if S_STATIC = 1 then echo "checked"%> />加速但不修改链接
    <input type="radio" name="static" onClick="set_static(2)" <%if S_STATIC = 2 then echo "checked"%> />加速并修改链接
    </td>
  </tr>
  <tr>
    <td>清除静态文件：</td>
    <td>
    <input type="text" id="clear_file_prefix" style="width:80px;" />&nbsp;&nbsp;
    <input type="button" onClick="clear_static()" value="立即清除" />&nbsp;&nbsp;全部清除请填0
    </td>
  </tr>
</table>
</div>
<div class="here">邮件通知设置</div>
<div class="block">
<form name="form_set_sentmail">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="30%">启用邮件通知功能：</td>
		<td>
			<%dim sentmail:sentmail = int(get_config("sentmail"))%>
			<input type="radio" name="sentmail" onClick="set_sentmail(1)" <%if sentmail = 1 then echo "checked"%> />开启
			<input type="radio" name="sentmail" onClick="set_sentmail(0)" <%if sentmail = 0 then echo "checked"%> />关闭
		</td>
	</tr>
	<tr>
		<td>发件SMTP服务器：</td>
		<td><input name="smtp" type="text" class="text" maxlength="200" value="<%=get_config("sentmail_smtp")%>" /></td>
	</tr>
	<tr>
		<td>发件邮箱：</td>
		<td><input name="send" type="text" class="text" maxlength="200" value="<%=get_config("sentmail_send")%>" /></td>
	</tr>
	<tr>
		<td>发件邮箱密码：</td>
		<td><input name="password" type="password" class="text" maxlength="200" value="<%=get_config("sentmail_password")%>" /></td>
	</tr>
	<tr>
		<td>收件邮箱：</td>
		<td><input name="receive" type="text" class="text" maxlength="200" value="<%=get_config("sentmail_receive")%>" /></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="button" onClick="save_sentmail()" value="保存" /></td>
	</tr>
</table>
</form>
</div>
<div class="here">使用说明</div>
<div class="block content">
  1.启用静态模式可以提高网页的加载速度，减轻服务器的负担。开启静态模式后，后台更新之后，必须清除静态文件才能在前台显示更新内容。<br />
  2.清除静态文件功能的文本框，可以填写要删除的文件前缀，例如“article”，留空不填时删除所有静态文件。<br />
  3.邮件通知功能的目前的主要用途是：当有访客在网站提交留言或调查问卷时，系统自动使用发件邮箱发送邮件通知管理员。<br />
  4.SMTP服务器地址一般形如smtp.163.com，但并不是所有SMTP服务器地址都是这种形式，163邮箱经多次测试发送邮件正常，建议使用。
</div>
<iframe id="deal" name="deal"></iframe>
<script language="javascript">
function set_about_filter(val)
{
	ajax("post","deal.asp","cmd=set_about_filter&val="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function set_about_len(val)
{
	ajax("post","deal.asp","cmd=set_about_len&val="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function set_list_len(con_name,val)
{
	ajax("post","deal.asp","cmd=set_list_len&con_name="+con_name+"&val="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function set_reduce_img(wh,val)
{
	ajax("post","deal.asp","cmd=set_reduce_img&con_name=reduce_img_"+wh+"&val="+val,
	function(data)
	{
	   if(data == 1) result();
	});
}
function set_session_method(val)
{
	ajax("post","deal.asp","cmd=set_session_method&val="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function set_compile(val)
{
	ajax("post","deal.asp","cmd=set_compile&val="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function set_static(val)
{
	ajax("post","deal.asp","cmd=set_static&val="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function clear_static()
{
	prefix = document.getElementById("clear_file_prefix").value;
	if(prefix != "")
	{
		ajax("post","deal.asp","cmd=clear_static&prefix="+prefix,
		function(data)
		{
			if(data == 1) result();
		});
	}
}
function set_sentmail(val)
{
	ajax("post","deal.asp","cmd=set_sentmail&val="+val,
	function(data)
	{
		if(data == 1) result();
	});
}
function save_sentmail()
{
	var smtp = document.form_set_sentmail.smtp.value;
	var send = document.form_set_sentmail.send.value;
	var password = document.form_set_sentmail.password.value;
	var receive = document.form_set_sentmail.receive.value;
	ajax("post","deal.asp","cmd=save_sentmail&smtp="+smtp+"&send="+send+"&password="+password+"&receive="+receive,
	function(data)
	{
		if(data == 1) result();
	});
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>