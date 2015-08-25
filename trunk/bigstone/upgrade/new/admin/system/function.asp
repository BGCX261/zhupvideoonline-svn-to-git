<%
'修改配置文件
function edit_config(n,c,v)
  dim path:path = server.mappath(S_ROOT&"include/config.asp")
  '读取配置文件
  dim stm:set stm = server.createobject("adodb.stream")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.loadfromfile path
	dim config:config = .readtext
	.close
  end with
  '替换文本并删除配置文件
  config = replace(config,n&" = "&c,n&" = "&v)
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  fso.deletefile(path)
  set fso = nothing
  '写入配置文件
  with stm
	.open
	.writetext config
	.savetofile path,2
	.close
  end with
  set stm = nothing
end function

'上传文件
function upload(byval path,byval fil)
  dim fso,folder,length,data,enter,start,divider,end_,stm,strm
  folder = server.mappath(path)
  set fso = server.createobject("scripting.filesystemobject")
  if not fso.folderexists(folder) then fso.createfolder(folder)
  set fso = nothing
  path = folder&"\"&fil
  length = request.totalbytes
  if length > 0 then
	data = request.binaryread(length)
	enter = chrb(13)&chrb(10)
	start = instrb(data,enter&enter) + 3
	divider = leftb(data,instrb(data,enter) - 1)
	end_ = instrb(start,data,divider) - start
	set stm = createobject("adodb.stream")
	with stm
	  .type = 1
	  .mode = 3
	  .open
	  .write data
	end with
	set  strm = createobject("adodb.stream")
	with strm
	  .type = 1
	  .mode = 3
	  .open
	  stm.position = start
	  stm.copyto strm,end_
	  .savetofile path,2
	  .close
	end with
	set strm = nothing
	set stm = nothing
  end if
end function

'生成缩略图
function reduce(img_dir,img_fil,max_width,max_height)
  on error resume next
  dim old_path:old_path = server.mappath(img_dir&img_fil)
  dim new_path:new_path = server.mappath(img_dir&"reduce_"&img_fil)  
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  if fso.fileexists(old_path) and not fso.fileexists(new_path) then fso.copyfile old_path,new_path
  dim jpeg:set jpeg = server.createobject("persits.jpeg")
  jpeg.open new_path
  if not err then
	if jpeg.width > max_width and jpeg.height > max_height then
	  if jpeg.width / jpeg.height > max_width / max_height then
		jpeg.height = max_width * jpeg.height / jpeg.width
		jpeg.width = max_width
	  else
		jpeg.width = max_height * jpeg.width / jpeg.height
		jpeg.height = max_height
	  end if
	elseif jpeg.width > max_width then
	  jpeg.height = max_width * jpeg.height / jpeg.width
	  jpeg.width = max_width
	elseif jpeg.height > max_height then
	  jpeg.width = max_height * jpeg.width / jpeg.height
	  jpeg.height = max_height
	else
	  jpeg.width = jpeg.width
	  jpeg.height = jpeg.height
	end if
	jpeg.save new_path
	set jpeg = nothing
  else
	err.clear
  end if
  set fso = nothing
end function

'获取菜单配置
function get_nav_config(val)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_type = 'nav_"&val&"' order by con_id asc")
  dim record_count:record_count = db.get_record_count()
  dim i,arr()
  if record_count > 0 then
	redim preserve arr(record_count - 1,1)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"con_name")
	  arr(i,0) = right(arr(i,0),len(arr(i,0)) - 4)
	  arr(i,1) = db.sql_result(i,"con_value")
	next
  else
	redim preserve arr(-1)
  end if
  get_nav_config = arr
end function

'获取菜单
function get_nav(men_type,men_show)
  call db.reset()
  if men_show = 1 then
	call db.sql_query("select * from "&S_DB_PREFIX&"menu where men_type = '"&men_type&"' order by men_top desc,men_index desc,men_id asc")
  else
	call db.sql_query("select * from "&S_DB_PREFIX&"menu where men_show = 1 and men_type = '"&men_type&"' order by men_top desc,men_index desc,men_id asc")
  end if
  dim record_count:record_count = db.get_record_count()
  dim i,arr()
  if record_count > 0 then
	redim preserve arr(record_count - 1,5)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"men_id")
	  arr(i,1) = db.sql_result(i,"men_name")
	  arr(i,2) = db.sql_result(i,"men_url")
	  arr(i,3) = db.sql_result(i,"men_index")
	  arr(i,4) = db.sql_result(i,"men_top")
	  arr(i,5) = db.sql_result(i,"men_show")
	next
  else
	redim preserve arr(-1)
  end if
  get_nav = arr
end function

'获取所有分类option标签，a代表id下标，b代表分类名下标,c代表等级下标
function get_cat_option(arr,a,b,c,cat_id)
  dim i,j,selected
  for i = 0 to ubound(arr)
	if cat_id = arr(i,a) then
	  selected = "selected"
	else
	  selected = ""
	end if
	%><option value="<%=arr(i,a)%>" <%=selected%>><%for j = 1 to arr(i,c) - 1%>&nbsp;&nbsp;<%next%><%=arr(i,b)%></option><%
  next
end function

'设置并返回部分分类管理所需的option标签，参数a,b,c同上
function set_cat_option(arr,a,b,c,parent_id)
  dim i,j
  dim grade:grade = 0
  dim flag:flag = 0
  dim selected:selected = ""
  if parent_id <> 0 then
	for i = 0 to ubound(arr)
	  if parent_id = arr(i,a) then
		grade = arr(i,c)
		flag = 1
		selected = "selected"
	  else
		selected = ""
	  end if
	  if flag = 0 or selected <> "" then
	  %><option value="<%=arr(i,a)%>" <%=selected%>><%for j = 1 to arr(i,c) - 1%>&nbsp;&nbsp;<%next%><%=arr(i,b)%></option><%
	  else
		if grade >= arr(i,c) then flag = 0
		if i <> ubound(arr) then
		  if grade >= arr(i + 1,c) then flag = 0
		end if
	  end if
	next
  end if
end function

'对分类进行排序，调用了get_tree_order()，并对其返回结果进行处理，a代表ID下标，b代表父类ID下标，flag标志是否存在第二级分类
function set_cat_order(arr,a,b,flag)
  dim i,j,k,tree_order,arr1(),arr2(),arr3,arr4,arr5()
  dim u1:u1 = ubound(arr)
  dim u2:u2 = ubound(arr,2)
  if u1 > -1 and flag = 1 then
	redim preserve arr1(u1)
	redim preserve arr2(u1)
	for i = 0 to u1
	  arr1(i) = arr(i,a)
	  arr2(i) = arr(i,b)
	next
	if u1 > -1 then
	  tree_order = split(get_tree_order(arr1,arr2),"{v}")
	  arr3 = split(tree_order(0),"|")
	  arr4 = split(tree_order(1),"|")
	  redim preserve arr5(u1,u2 + 1)
	  for i = 1 to ubound(arr3) - 1
		for j = 0 to u1
		  if arr(j,0) = arr3(i) then
		    for k = 0 to u2
			  arr5(i - 1,k) = arr(j,k)
			next
			arr5(i - 1,u2 + 1) = arr4(i)
		  end if
		next
	  next
	end if
  elseif u1 > -1 and flag = 0 then
    redim preserve arr5(u1,u2 + 1)
	for i = 0 to u1
	  for k = 0 to u2
		arr5(i,k) = arr(i,k)
	  next
	  arr5(i,u2 + 1) = 1
	next
  else
	redim preserve arr5(-1)
  end if
  set_cat_order = arr5
end function

'对无序多级分类进行排序，返回一个记录排序信息的特殊字符串
function get_tree_order(id,parent)
  dim i,id_arr,parent_arr,cat_str
  dim j:j = 0
  dim k:k = ubound(parent) + 1
  dim id_str()
  dim parent_str()
  dim grade_str:grade_str = "|"
  '将分类划分为不同等级
  do while k > 0
	redim preserve id_str(j)
	redim preserve parent_str(j)
	id_str(j) = "|"
	parent_str(j) = "|"
	for i = 0 to ubound(parent)
	  if j = 0 then
		if parent(i) = 0 then
		  id_str(j) = id_str(j) & id(i) & "|"
		  parent_str(j) = parent_str(j) & parent(i) & "|"
		  k = k - 1
		end if
	  else
		if instr(id_str(j - 1),"|"&parent(i)&"|") > 0 then
		  id_str(j) = id_str(j) & id(i) & "|"
		  parent_str(j) = parent_str(j) & parent(i) & "|"
		  k = k - 1
		end if
	  end if
	next
	j = j + 1
  loop
  '将子级分类倒序排列
  for i = 1 to ubound(id_str)
	dim str1:str1 = ""
	dim str2:str2 = ""
	id_arr = split(id_str(i),"|")
	parent_arr = split(parent_str(i),"|")
	for j = 1 to ubound(id_arr)
	  str1 = id_arr(j) & "|" & str1
	  str2 = parent_arr(j) & "|" & str2
	next
	id_str(i) = str1
	parent_str(i) = str2
  next
  '将子级分类插入父级分类后面
  if ubound(id_str) > 0 then cat_str = id_str(0)
  for i = 1 to ubound(id_str)
	id_arr = split(id_str(i),"|")
	parent_arr = split(parent_str(i),"|")
	for j = 1 to ubound(parent_arr) - 1
	  if instr(cat_str,"|"&parent_arr(j)&"|") > 0 then
		cat_str = replace(cat_str,"|"&parent_arr(j)&"|","|"&parent_arr(j)&"|"&id_arr(j)&"|")
	  end if
	next
  next
  '获取等级信息
  dim arr:arr = split(cat_str,"|")
  for i = 1 to ubound(arr) - 1
   for j = 0 to ubound(id_str)
	 if instr(id_str(j),"|"&arr(i)&"|") > 0 then grade_str = grade_str & (j + 1) & "|"
   next
  next
  get_tree_order = cat_str & "{v}" & grade_str
end function

'字节转换为字符串
function b2s(bytes)
  dim str,i,this_c,next_c
  str = "" 
  bytes = cstr(bytes) 
  for i = 1 to lenb(bytes)
	this_c = ascb(midb(bytes, i, 1))
	if this_c < &h80 then
	  str = str & chr(this_c)
	else
	  next_c = ascb(midb(bytes, i+1, 1))
	  str = str & chr(clng(this_c) * &h100 + cint(next_c))
	  i = i + 1
	end if
  next
  b2s = str
end function

'提取HTML文本中的图片
function get_img_in_text(text)
  dim reg:set reg = new regexp
  reg.ignorecase = true
  reg.global = true
  reg.pattern = "<img.*?>"
  dim mat
  dim matches:set matches = reg.execute(text)
  for each mat in matches
	img = mat.value
	exit for
  next
  set reg = nothing
  set matches = nothing
  get_img_in_text = img
end function

'获取模块名
function get_module_name(val)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"modules where mod_type='"&val&"'")
  if db.get_record_count() > 0 then
	dim arr:arr = split(db.sql_result(0,"mod_config"),"|")
	get_module_name = arr(0)
  else
	get_module_name = false
  end if
end function

'获取模块配置值
function get_module_config(val)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"modules where mod_type='"&val&"' order by mod_index desc,mod_id asc")
  dim record_count:record_count = db.get_record_count()
  dim i,arr()
  if record_count > 0 then
	redim preserve arr(record_count - 1,1)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"mod_id")
	  arr(i,1) = db.sql_result(i,"mod_config")
	next
  else
	redim preserve arr(-1)
  end if
  get_module_config = arr
end function

'翻页链接
function page_link(fil,page,page_sum,field_,search)
dim link:link = fil & "?from=" & page & "&module=" & module
if field_ <> "" then link = link & "&field_=" & field_
if search <> "" then link = link & "&search=" & search
link = link & "&page="
%>
<div class="page_link"><div class="in">
  <span><%=L_ALL&page_sum&L_PAGES%></span>
  <span><%=L_NO&page&L_PAGE%></span>
  <%if page_sum <> 1 then%>
    <a href="<%=link&"1"%>"><%=L_FIRST_PAGE%></a>
    <%if page-1 > 0 then%><a href="<%=link&(page-1)%>"><%=L_PREVIOUS_PAGE%></a><%end if%>
    <%if page-4 > 0 then%><a href="<%=link&(page-4)%>">[<%=page-4%>]</a><%end if%>
    <%if page-3 > 0 then%><a href="<%=link&(page-3)%>">[<%=page-3%>]</a><%end if%>
    <%if page-2 > 0 then%><a href="<%=link&(page-2)%>">[<%=page-2%>]</a><%end if%>
    <%if page-1 > 0 then%><a href="<%=link&(page-1)%>">[<%=page-1%>]</a><%end if%>
    <a href="<%=link&page%>" style="color:#C00;font-weight:bold;">[<%=page%>]</a>
    <%if page+1 <= page_sum then%><a href="<%=link&(page+1)%>">[<%=page+1%>]</a><%end if%>
    <%if page+2 <= page_sum then%><a href="<%=link&(page+2)%>">[<%=page+2%>]</a><%end if%>
    <%if page+3 <= page_sum then%><a href="<%=link&(page+3)%>">[<%=page+3%>]</a><%end if%>
    <%if page+4 <= page_sum then%><a href="<%=link&(page+4)%>">[<%=page+4%>]</a><%end if%>
    <%if page+1 <= page_sum then%><a href="<%=link&(page+1)%>"><%=L_NEXT_PAGE%></a><%end if%>
    <a href="<%=link&page_sum%>"><%=L_LAST_PAGE%></a>
  <%end if%>
  <form action="<%=fil%>" method="get">
    <input type="hidden" name="field_" value="<%=field_%>" />
    <input type="hidden" name="search" value="<%=search%>" />
    <input type="text" name="page" style="width:30px;" value="<%=page%>" />
    <input type="submit" value="<%=L_JUMP%>"/>
  </form>
</div></div>
<%
end function

'新秀
%>