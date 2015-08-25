<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim mod_name:mod_name = get_module_name(module)
dim rank:rank = get_module_config(module&"_form")
%>
<div class="here">添加<%=mod_name%></div>
<div class="block">
  <form name="form_add_art" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="add_art" />
    <input name="module" type="hidden" value="<%=module%>" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	  for i = 0 to ubound(rank)
		call set_form_row(rank(i,1),"")
	  next
	  %>
      <tr>
        <td>SEO：</td>
        <td>
          <input id="seo_but" type="button" onClick="set_seo('art')" value="进行优化" />
          <div style="display:none;">
            <input name="art_keywords" type="text" class="text" maxlength="200" value="<%=get_config("sit_keywords")%>" />
            <textarea class="area" name="art_description"><%=get_config("sit_description")%></textarea>
          </div>
        </td>
      </tr>
      <tr>
        <td class="button" colspan="30">
          <input type="submit" onClick="do_submit('form_add_art')" value="提交" />
        </td>
      </tr>
    </table>
  </form>
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="upl_img">
  <div class="head">
    <div class="title">上传图片</div>
    <div class="close" onClick="hid_box('upl_img')">关闭</div>
  </div>
  <input name="upl_img_dir" type="hidden" value="<%=S_ROOT%>images/article/<%=get_date()%>/" />
  <form name="upl_img_form" method="post" enctype="multipart/form-data" target="deal">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><input name="upl_img_path" type="file" onChange="do_upload('upl_img')" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="box" id="upl_fil">
  <div class="head">
    <div class="title">上传文件</div>
    <div class="close" onClick="hid_box('upl_fil')">关闭</div>
  </div>
  <input name="upl_fil_dir" type="hidden" value="<%=S_ROOT%>download/" />
  <form name="upl_fil_form" method="post" enctype="multipart/form-data" target="deal">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><input name="upl_fil_path" type="file" onChange="do_upload('upl_fil')" /></td>
      </tr>
    </table>
  </form>
</div>
<div class="box" id="web_fil">
  <div class="head">
    <div class="title">网络文件</div>
    <div class="close" onClick="hid_box('web_fil')">关闭</div>
  </div>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr><td>
      <input name="fil_url" type="text" value="http://" />
      <input type="button" onClick="set_web_url()" value="确定" />
    </td></tr>
  </table>
</div>
<div class="box" id="seo">
  <div class="head">
    <div class="title">进行简单优化</div>
    <div class="close" onClick="hid_box('seo')">关闭</div>
  </div>
  <form name="" method="post" action="">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>页面关键字：</td>
        <td style="text-align:left;">
          <input name="seo_key" type="text" style="width:300px;" maxlength="200" value="" />
        </td>
      </tr>
      <tr>
        <td>页面描述：</td>
        <td style="text-align:left;">
          <textarea class="area" name="seo_des"></textarea>
        </td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onClick="get_seo('art')" value="确定" /></td>
      </tr>
    </table>
  </form>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
function do_submit(val)
{
	switch(val)
	{
		case "form_add_art":break;
	}
	interval = setInterval("get_result()",1000);
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>