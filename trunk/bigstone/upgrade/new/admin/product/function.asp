<%
function get_art_sheet(search,field_,page,page_sum)
  dim sql,arr(),fie_arr
  if search = "" then
	sql = "select * from "&S_DB_PREFIX&"article where art_type = '"&module&"' order by art_top desc,art_index desc, art_id desc"
  else
	sql = "select * from "&S_DB_PREFIX&"article where art_type = '"&module&"' and "&field_&" like '%"&search&"%' order by art_top desc,art_index desc, art_id desc"
  end if
  fie_arr = get_module_fields(module&"_sheet")
  call db.reset()
  call db.sql_query(sql)
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	dim page_size:page_size = 10
	call db.set_page_size(page_size)
	if (record_count mod page_size) > 0 then
	  page_sum = int(record_count / page_size) + 1
	else
	  page_sum = int(record_count / page_size)
	end if
	page = num_bound(1,page_sum,page)
	call db.set_absolute_page(page)
	record_count = record_count - (page - 1) * page_size
	if record_count < page_size then
	  page_size = record_count
	end if
	call db.set_real_page_size(page_size)
	redim preserve arr(page_size - 1,ubound(fie_arr) + 1)
	dim i:for i = 0 to page_size - 1
	  arr(i,0) = db.sql_result(i,"art_id")
	  dim j:for j = 0 to ubound(fie_arr)
		if left(fie_arr(j),3) = "art" then
		  arr(i,j+1) = db.sql_result(i,fie_arr(j))
		else
		  arr(i,j+1) = get_attribute(db.sql_result(i,"art_attributes"),fie_arr(j))
		end if
	  next
	next
  else
	redim preserve arr(-1)
	page_sum = 1
	page = 1
  end if
  get_art_sheet = arr
end function

function get_module_fields(val)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"modules where mod_type='"&val&"' order by mod_index desc,mod_id asc")
  dim record_count:record_count = db.get_record_count()
  dim i,arr(),arr2
  if record_count > 0 then
	redim preserve arr(record_count - 1)
	for i = 0 to record_count - 1
	  arr2 = split(db.sql_result(i,"mod_config"),"|")
	  arr(i) = arr2(0)
	next
  else
	redim preserve arr(-1)
  end if
  get_module_fields = arr
end function

function get_cat_sheet()
  dim i,arr()
  dim flag:flag = 0
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"category where cat_type = '"&module&"' order by cat_top desc,cat_index desc,cat_id asc")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	redim preserve arr(record_count - 1,5)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"cat_id")
	  arr(i,1) = db.sql_result(i,"cat_parent_id")
	  arr(i,2) = db.sql_result(i,"cat_name")
	  arr(i,3) = db.sql_result(i,"cat_index")
	  arr(i,4) = db.sql_result(i,"cat_top")
	  arr(i,5) = db.sql_result(i,"cat_show")
	  if arr(i,1) <> 0 then flag = 1
	next
  else
	redim preserve arr(-1)
  end if
  if ubound(arr) >= 0 then get_cat_sheet = set_cat_order(arr,0,1,flag) else get_cat_sheet = arr
end function

function get_cat_name(val)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"category where cat_type = '"&module&"' and cat_id = "&val)
  if db.get_record_count() > 0 then
	get_cat_name = db.sql_result(0,"cat_name")
  else
    get_cat_name = L_NONE
  end if
end function

function exist_child(val)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"category where cat_type = '"&module&"' and cat_parent_id = "&val)
  if db.get_record_count() > 0 then
	exist_child = 1
  else
    exist_child = 0
  end if
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function set_sheet_rank(rank,id,val)
  dim arr:arr = split(rank,"|")
  dim func:func = arr(ubound(arr)) & "(id,val)"
  eval(func)
end function

function set_form_row(rank,val)
  dim arr:arr = split(rank,"|")
  dim func:func = arr(3) & "(arr(1),arr(2),val)"
  eval(func)
end function

function set_sheet_head(rank)
  dim arr:arr = split(rank,"|")
  %>
  <td width="<%=arr(2)%>"><b><%=arr(1)%></b></td>
  <%
end function

function befort_sheet_img(byref id,byref val)
%>
<td class="pic">
  <div class="unit">
    <table cellspacing="0" cellpadding="0">
      <tr><td>
        <img src="<%=S_ROOT&val%>" onload="picresize(this,50,50)"/>
      </td></tr>
    </table>
  </div>
</td>
<%
end function

function befort_sheet_word(byref id,byref val)
%>
<td><%=val%></td>
<%
end function

function befort_sheet_cat(byref id,byref val)
dim cat_name:cat_name = get_data("category",val,"cat_name")
%>
<td><%if cat_name <> false then%><%=cat_name%><%else%>无<%end if%></td>
<%
end function

function befort_sheet_link(byref id,byref val)
%>
<td>?<%=module%>_<%=id%>.<%=S_SUFFIX%></td>
<%
end function

function befort_sheet_index(byref id,byref val)
%>
<td onClick="show_edit('index_<%=id%>')">
<span id="index_<%=id%>_1"><%=val%><img src="../system/images/pencil.gif"></span>
<span id="index_<%=id%>_2" style="display:none;"><input type="text" id="index_<%=id%>" value="<%=val%>" style="width:30px;" onBlur="set_index(<%=id%>,this.value,'art')" /></span>
</td>
<%
end function

function befort_sheet_top(byref id,byref val)
%>
<td><input style="border:0;" type="checkbox" <%if val = 1 then%>checked<%end if%> onClick="set_top(<%=id%>,this.checked,'art')" /></td>
<%
end function

function befort_sheet_show(byref id,byref val)
%>
<td><input style="border:0;" type="checkbox" <%if val = 1 then%>checked<%end if%> onClick="set_show(<%=id%>,this.checked,'art')" /></td>
<%
end function

function befort_form_text(byref nam,byref word,byref val)
%>
<tr>
  <td><%=word%>：</td>
  <td><input name="<%=nam%>" type="text" class="text" maxlength="200" value="<%=val%>" /></td>
</tr>
<%
end function

function befort_form_cat(byref nam,byref word,byref val)
if val = "" then val = 0
%>
<tr>
  <td><%=word%>：</td>
  <td>
  <%dim arr:arr = get_cat_sheet()%>
  <select name="<%=nam%>">
    <option value="0">请选择</option>
    <%=get_cat_option(arr,0,2,6,val)%>
  </select>
  </td>
</tr>
<%
end function

function befort_form_upload_img(byref nam,byref word,byref val)
%>
<!--[if IE 6]>
<tr>
  <td colspan="30"><span class="red">重要提示：</span>
  您的浏览器版本较低（IE6或以IE6为内核的其它浏览器），不能上传中文名的图片，建议使用IE7、IE8、IE9、火狐、谷歌等浏览器
  </td>
</tr>
<![endif]-->
<tr>
  <td><%=word%>：</td>
  <td>
  <%if val = "" then%>
  <span id="get_pic"></span>
  <%else%>
  <span id="get_pic"><img src="<%=S_ROOT&val%>" height="100px" /></span>
  <%end if%>
  <input type="button" onClick="show_box('upl_img',320,90)" value="上传图片" />
  <%if val = "" then%>
  <input name="<%=nam%>" type="hidden" maxlength="250" value="" />
  <%else%>
  <input name="<%=nam%>" type="hidden" maxlength="250" value="<%=S_ROOT&val%>" />
  <%end if%>
  </td>
</tr>
<%
end function

function befort_form_upload_file(byref nam,byref word,byref val)
%>
<tr>
  <td><%=word%>：</td>
  <td>
  <input type="button" onClick="show_box('upl_fil',320,90)" value="上传文件" />
  <input type="button" onClick="show_box('web_fil',320,90)" value="网络文件" />
  <%if val = "" then%>
  <input name="<%=nam%>" type="hidden" maxlength="250" value="" />
  <span id="get_fil"></span>
  <%else%>
  <input name="<%=nam%>" type="hidden" maxlength="250" value="<%=S_ROOT&val%>" />
  <span id="get_fil"><%if instr(replace(val,"http;//","http://"),"http://") > 0 then echo val else echo S_ROOT&val%></span>
  <%end if%>
  </td>
</tr>
<%
end function

function befort_form_editor(byref nam,byref word,byref val)
%>
<tr>
  <td><%=word%>：</td>
  <td><div class="editor"><%call interface("admin_editor",val)%></div></td>
</tr>
<%
end function


'新秀
%>