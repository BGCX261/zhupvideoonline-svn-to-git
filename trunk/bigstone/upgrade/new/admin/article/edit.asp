<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim id:id = gets("id","")
dim page:page = gets("page","")
dim mod_name:mod_name = get_module_name(module)
dim rank:rank = get_module_config(module&"_form")
call db.reset()
call db.sql_query("select * from "&S_DB_PREFIX&"article where art_id = "&id)
if db.get_record_count() > 0 then
  dim art_id:art_id = db.sql_result(0,"art_id")
  dim art_type:art_type = db.sql_result(0,"art_type")
  dim art_cat_id:art_cat_id = db.sql_result(0,"art_cat_id")
  dim art_title:art_title = db.sql_result(0,"art_title")
  dim art_text:art_text = db.sql_result(0,"art_text")
  dim art_short_text:art_short_text = db.sql_result(0,"art_short_text")
  dim art_img:art_img = db.sql_result(0,"art_img")
  dim art_reduce_img:art_reduce_img = db.sql_result(0,"art_reduce_img")
  dim art_more_img:art_more_img = db.sql_result(0,"art_more_img")
  dim art_attributes:art_attributes = db.sql_result(0,"art_attributes")
  dim art_add_time:art_add_time = db.sql_result(0,"art_add_time")
  dim art_top:art_top = db.sql_result(0,"art_top")
  dim art_index:art_index = db.sql_result(0,"art_index")
  dim art_show:art_show = db.sql_result(0,"art_show")
  dim art_keywords:art_keywords = db.sql_result(0,"art_keywords")
  dim art_description:art_description = db.sql_result(0,"art_description")
  dim art_hits:art_hits = db.sql_result(0,"art_hits")
end if
%>
<div class="here">编辑<%=mod_name%>信息</div>
<div class="block">
  <form name="form_edit_art" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_art" />
    <input name="module" type="hidden" value="<%=module%>" />
    <input name="id" type="hidden" value="<%=id%>" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <%
	  for i = 0 to ubound(rank)
	    dim arr:arr = split(rank(i,1),"|")
		if left(arr(0),3) = "art" then
		  call set_form_row(rank(i,1),eval(arr(0)))
		else
		  call set_form_row(rank(i,1),get_attribute(art_attributes,arr(0)))
		end if
	  next
	  %>
      <tr>
        <td>SEO：</td>
        <td>
          <input id="seo_but" type="button" onClick="set_seo('art')" value="进行优化" />
          <div style="display:none;">
            <input name="art_keywords" type="text" class="text" maxlength="200" value="<%=art_keywords%>" />
            <textarea class="area" name="art_description"><%=art_description%></textarea>
          </div>
        </td>
      </tr>
      <tr>
        <td class="button" colspan="30">
          <input type="submit" onClick="do_submit('form_edit_art')" value="提交" />
          <input type="button" onClick="location.replace(document.referrer)" value="返回" />
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
		case 'form_edit_art':break;
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