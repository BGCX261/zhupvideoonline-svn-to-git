<!-- #include file="../system/common.asp" -->
<%
dim func:func = post("cmd","") & "()"
eval(func)

function create_page()
  dim old_page:old_page = post("old_page","")
  dim new_page:new_page = post("new_page","")
  dim new_name:new_name = post("new_name","")
  dim flag:flag = is_create_page(old_page,new_page,new_name)
  if flag = 1 then
	dim func:func = "create_" & old_page & "_page(new_page,new_name)"
	eval(func)
  else
    echo flag
  end if
end function

function is_create_page(old_page,new_page,new_name)
  new_page = replace(new_page,"?","")
  new_page = replace(new_page,".html","")
  select case new_page
	case "njb":new_page = ""
	case "index":new_page = ""
	case "about":new_page = ""
	case "product":new_page = ""
	case "article":new_page = ""
	case "recruit":new_page = ""
	case "download":new_page = ""
	case "message":new_page = ""
	case "search":new_page = ""
  end select
  if old_page <> "" and new_page <> "" and new_name <> "" then
	dim i,val
	dim flag:flag = 1
	for i = 1 to len(new_page)
	  val = mid(new_page,i,1)
	  if i = 1 then
		if asc(val) < asc("a") or asc(val) > asc("z") then
		  flag = 0
		  exit for
		end if
	  else
		if (asc(val) < asc("0") or asc(val) > asc("9")) and (asc(val) < asc("a") or asc(val) > asc("z")) then
		  flag = 0
		  exit for
		end if
	  end if
	next
	if flag = 1 then
	  call db.reset()
	  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name = 'new_page' and left(con_value,len('"&new_page&"')+3) = '"&new_page&"{v}'")
	  if db.get_record_count() = 0 then
		is_create_page = 1
	  else
	    is_create_page = 4
	  end if
	else
	  is_create_page = 3
	end if
  else
	is_create_page = 2
  end if
end function

function copy_page(old_path,old_word,new_path,new_word)
  set stm = server.createobject("adodb.str"&"eam")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.loadfromfile server.mappath(old_path)
	text = .readtext
	.close
  end with
  text = protect_replace(text,1)
  text = replace(text,old_word,new_word)
  text = protect_replace(text,0)
  set fso = server.createobject("scripting.filesystemobject")
  set fso = nothing
  with stm
	.open
	.writetext text
	.savetofile server.mappath(new_path),2
	.close
  end with
  set stm = nothing
end function

function protect_replace(text,val)
  if val = 1 then
	text = replace(text,"call set_head(""article"")","call set_head(""njb"")")
	text = replace(text,"article_model","njb_model")
	text = replace(text,"get_article()","get_njb()")
	text = replace(text,"get_article_name()","get_njb_name()")
	text = replace(text,"=""article""","=""njb""")
  else
	text = replace(text,"call set_head(""njb"")","call set_head(""article"")")
	text = replace(text,"njb_model","article_model")
	text = replace(text,"get_njb()","get_article()")
	text = replace(text,"get_njb_name()","get_article_name()")
	text = replace(text,"=""njb""","=""article""")
  end if
  protect_replace = text
end function

function copy_code(path,copy_start,copy_end,flag,old_word,new_word)
  dim text
  set stm = server.createobject("adodb.str"&"eam")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.loadfromfile server.mappath(path)
	text = .readtext
	.close
  end with
  dim a:a = instr(text,copy_start)
  dim str:str = mid(text,a,len(text)-a)
  dim b:b = instr(str,copy_end) + len(copy_end)
  str = protect_replace(mid(text,a,b),1)
  str = replace(str,old_word,new_word)
  str = protect_replace(str,0)
  text = replace(text,flag,str&chr(13)&chr(10)&flag)
  set fso = server.createobject("scripting.filesystemobject")
  set fso = nothing
  with stm
	.open
	.writetext text
	.savetofile server.mappath(path),2
	.close
  end with
  set stm = nothing
end function

function del_code(path,copy_start,copy_end)
  dim text
  set stm = server.createobject("adodb.str"&"eam")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.loadfromfile server.mappath(path)
	text = .readtext
	.close
  end with
  dim a:a = instr(text,copy_start)
  dim str:str = mid(text,a,len(text)-a)
  dim b:b = instr(str,copy_end) + len(copy_end)
  if instr(text,mid(text,a,b)&chr(13)&chr(10)) <> 0 then
	text = replace(text,mid(text,a,b)&chr(13)&chr(10),"")
  else
	text = replace(text,mid(text,a,b),"")
  end if
  set fso = server.createobject("scripting.filesystemobject")
  set fso = nothing
  with stm
	.open
	.writetext text
	.savetofile server.mappath(path),2
	.close
  end with
  set stm = nothing
end function

function reset_xml_name(new_page,new_name)
  dim sinty_xml:sinty_xml = server.mappath(S_TPL_PATH&"sinty.xml")
  dim xml:set xml = server.createobject("microsoft.xmldom")
  call xml.load(sinty_xml)
  call xml.selectsinglenode("site").selectsinglenode(new_page).setattribute("name",new_name)
  call xml.save(sinty_xml)
  set xml = nothing
end function

function create_about_page(byref new_page,byref new_name)
  '复制和修改前台文件
  call copy_page(S_ROOT&"index/about.asp","about",S_ROOT&"index/"&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"about.asp","about",S_TPL_PATH&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"module/about_menu.asp","about",S_TPL_PATH&"module/"&new_page&"_menu.asp",new_page)
  call copy_page(S_TPL_PATH&"module/about_main.asp","about",S_TPL_PATH&"module/"&new_page&"_main.asp",new_page)
  call copy_code(S_TPL_PATH&"sinty.xml","<about","</about>","</site>","about",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_about_menu()","end function","'new function","about",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_about_main()","end function","'new function","about",new_page)
  call reset_xml_name(new_page,new_name)
  '复制后台目录
  set fso = createobject("scripting.filesystemobject")
  fso.copyfolder server.mappath(S_ROOT&"admin/about"),server.mappath(S_ROOT&"admin/"&new_page)
  set fso = nothing
  '修改数据库
  call db.reset()
  db.sql_query("insert into modules (mod_type,mod_action,mod_config) values ('"&new_page&"','global','"&new_name&"|"&new_page&"|0|0')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_title|标题|auto|befort_sheet_word','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_id|页面链接|150px|befort_sheet_link','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_index|排序|45px|befort_sheet_index','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_top|置顶|40px|befort_sheet_top','7')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_show|显示|40px|befort_sheet_show','6')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_title|art_title|文章标题|befort_form_text|deal_form_text','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_text|editor|文章正文|befort_form_editor|deal_form_editor','9')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin','"&new_name&"','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','"&new_name&"列表','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','添加"&new_name&"','main.asp?dir="&new_page&"&fil=add')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('header','"&new_name&"','?"&new_page&".html')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('default','new_page','"&new_page&"{v}"&new_name&"{v}about')")
  db.sql_query("insert into config (con_name,con_value) values ('title','"&new_page&"{v}"&new_name&"{v}10')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('nav_stage','nav_"&new_page&"','"&new_name&"')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('nav_admin','nav_admin_"&new_page&"','"&new_name&"')")
  call set_session("result",1,"")
  echo 1
end function

function create_product_page(byref new_page,byref new_name)
  '复制和修改前台文件
  call copy_page(S_ROOT&"index/product.asp","product",S_ROOT&"index/"&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"product.asp","product",S_TPL_PATH&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"module/product_tree.asp","product",S_TPL_PATH&"module/"&new_page&"_tree.asp",new_page)
  call copy_page(S_TPL_PATH&"module/product_main.asp","product",S_TPL_PATH&"module/"&new_page&"_main.asp",new_page)
  call copy_code(S_TPL_PATH&"sinty.xml","<product","</product>","</site>","product",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_product_tree()","end function","'new function","product",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_product_main()","end function","'new function","product",new_page)
  call reset_xml_name(new_page,new_name)
  '复制后台目录
  set fso = createobject("scripting.filesystemobject")
  fso.copyfolder server.mappath(S_ROOT&"admin/product"),server.mappath(S_ROOT&"admin/"&new_page)
  set fso = nothing
  '修改数据库
  call db.reset()
  db.sql_query("insert into modules (mod_type,mod_action,mod_config) values ('"&new_page&"','global','"&new_name&"|none|0|0')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_reduce_img|图片|auto|befort_sheet_img','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_title|标题|auto|befort_sheet_word','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_index|排序|45px|befort_sheet_index','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_top|置顶|40px|befort_sheet_top','7')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_show|显示|40px|befort_sheet_show','6')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_title|art_title|产品标题|befort_form_text|deal_form_text','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','att_number|att_number|产品编号|befort_form_text|deal_form_att','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_cat_id|art_cat_id|产品分类|befort_form_cat|deal_form_text','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_img|pic_path|上传图片|befort_form_upload_img|deal_form_img','7')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','att_price|att_price|产品价格|befort_form_text|deal_form_att','6')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_text|editor|产品描述|befort_form_editor|deal_form_editor','5')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin','"&new_name&"','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','"&new_name&"列表','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','添加"&new_name&"','main.asp?dir="&new_page&"&fil=add')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','"&new_name&"分类','main.asp?dir="&new_page&"&fil=cat')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('header','"&new_name&"','?"&new_page&".html')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('default','new_page','"&new_page&"{v}"&new_name&"{v}product')")
  db.sql_query("insert into config (con_name,con_value) values ('title','"&new_page&"{v}"&new_name&"{v}10')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('nav_admin','nav_admin_"&new_page&"','"&new_name&"')")
  call set_session("result",1,"")
  echo 1
end function

function create_article_page(byref new_page,byref new_name)
  '复制和修改前台文件
  call copy_page(S_ROOT&"index/article.asp","article",S_ROOT&"index/"&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"article.asp","article",S_TPL_PATH&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"module/article_tree.asp","article",S_TPL_PATH&"module/"&new_page&"_tree.asp",new_page)
  call copy_page(S_TPL_PATH&"module/article_main.asp","article",S_TPL_PATH&"module/"&new_page&"_main.asp",new_page)
  call copy_code(S_TPL_PATH&"sinty.xml","<article","</article>","</site>","article",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_article_tree()","end function","'new function","article",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_article_main()","end function","'new function","article",new_page)
  call reset_xml_name(new_page,new_name)
  '复制后台目录
  set fso = createobject("scripting.filesystemobject")
  fso.copyfolder server.mappath(S_ROOT&"admin/article"),server.mappath(S_ROOT&"admin/"&new_page)
  set fso = nothing
  '修改数据库
  call db.reset()
  db.sql_query("insert into modules (mod_type,mod_action,mod_config) values ('"&new_page&"','global','"&new_name&"|none|0|0')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_cat_id|分类|100px|befort_sheet_cat','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_title|标题|auto|befort_sheet_word','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_index|排序|45px|befort_sheet_index','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_top|置顶|40px|befort_sheet_top','7')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_show|显示|40px|befort_sheet_show','6')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_title|art_title|文章标题|befort_form_text|deal_form_text','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_cat_id|art_cat_id|文章分类|befort_form_cat|deal_form_text','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','att_author|att_author|作者/来源|befort_form_text|deal_form_att','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_text|editor|产品描述|befort_form_editor|deal_form_editor','7')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin','"&new_name&"','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','"&new_name&"列表','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','添加"&new_name&"','main.asp?dir="&new_page&"&fil=add')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','"&new_name&"分类','main.asp?dir="&new_page&"&fil=cat')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('header','"&new_name&"','?"&new_page&".html')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('default','new_page','"&new_page&"{v}"&new_name&"{v}article')")
  db.sql_query("insert into config (con_name,con_value) values ('title','"&new_page&"{v}"&new_name&"{v}10')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('nav_admin','nav_admin_"&new_page&"','"&new_name&"')")
  call set_session("result",1,"")
  echo 1
end function

function create_recruit_page(byref new_page,byref new_name)
  '复制和修改前台文件
  call copy_page(S_ROOT&"index/recruit.asp","recruit",S_ROOT&"index/"&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"recruit.asp","recruit",S_TPL_PATH&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"module/recruit_main.asp","recruit",S_TPL_PATH&"module/"&new_page&"_main.asp",new_page)
  call copy_code(S_TPL_PATH&"sinty.xml","<recruit","</recruit>","</site>","recruit",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_recruit_main()","end function","'new function","recruit",new_page)
  call reset_xml_name(new_page,new_name)
  '复制后台目录
  set fso = createobject("scripting.filesystemobject")
  fso.copyfolder server.mappath(S_ROOT&"admin/recruit"),server.mappath(S_ROOT&"admin/"&new_page)
  set fso = nothing
  '修改数据库
  call db.reset()
  db.sql_query("insert into modules (mod_type,mod_action,mod_config) values ('"&new_page&"','global','"&new_name&"|none|0|0')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_title|标题|auto|befort_sheet_word','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_index|排序|45px|befort_sheet_index','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_top|置顶|40px|befort_sheet_top','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_show|显示|40px|befort_sheet_show','7')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_title|art_title|文章标题|befort_form_text|deal_form_text','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_text|editor|文章正文|befort_form_editor|deal_form_editor','9')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin','"&new_name&"','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','"&new_name&"列表','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','添加"&new_name&"','main.asp?dir="&new_page&"&fil=add')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('header','"&new_name&"','?"&new_page&".html')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('default','new_page','"&new_page&"{v}"&new_name&"{v}recruit')")
  db.sql_query("insert into config (con_name,con_value) values ('title','"&new_page&"{v}"&new_name&"{v}10')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('nav_admin','nav_admin_"&new_page&"','"&new_name&"')")
  call set_session("result",1,"")
  echo 1
end function

function create_download_page(byref new_page,byref new_name)
  '复制和修改前台文件
  call copy_page(S_ROOT&"index/download.asp","download",S_ROOT&"index/"&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"download.asp","download",S_TPL_PATH&new_page&".asp",new_page)
  call copy_page(S_TPL_PATH&"module/download_main.asp","download",S_TPL_PATH&"module/"&new_page&"_main.asp",new_page)
  call copy_code(S_TPL_PATH&"sinty.xml","<download","</download>","</site>","download",new_page)
  call copy_code(S_ROOT&"index/module.func.asp","function module_download_main()","end function","'new function","download",new_page)
  call reset_xml_name(new_page,new_name)
  '复制后台目录
  set fso = createobject("scripting.filesystemobject")
  fso.copyfolder server.mappath(S_ROOT&"admin/download"),server.mappath(S_ROOT&"admin/"&new_page)
  set fso = nothing
  '修改数据库
  call db.reset()
  db.sql_query("insert into modules (mod_type,mod_action,mod_config) values ('"&new_page&"','global','"&new_name&"|none|0|0')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_title|标题|auto|befort_sheet_word','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_index|排序|45px|befort_sheet_index','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_top|置顶|40px|befort_sheet_top','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_sheet','main','art_show|显示|40px|befort_sheet_show','7')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_title|art_title|文章标题|befort_form_text|deal_form_text','10')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','att_author|att_author|作者/来源|befort_form_text|deal_form_att','9')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','att_url|fil_path|上传文件|befort_form_upload_file|deal_form_file','8')")
  db.sql_query("insert into modules (mod_type,mod_action,mod_config,mod_index) values ('"&new_page&"_form','main','art_text|editor|文章正文|befort_form_editor|deal_form_editor','7')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin','"&new_name&"','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','"&new_name&"列表','main.asp?dir="&new_page&"&fil=sheet')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('admin_"&new_page&"','添加"&new_name&"','main.asp?dir="&new_page&"&fil=add')")
  db.sql_query("insert into menu (men_type,men_name,men_url) values ('header','"&new_name&"','?"&new_page&".html')")
  db.sql_query("insert into config (con_name,con_value) values ('new_page','"&new_page&"{v}"&new_name&"{v}download')")
  db.sql_query("insert into config (con_name,con_value) values ('title','"&new_page&"{v}"&new_name&"{v}10')")
  db.sql_query("insert into config (con_type,con_name,con_value) values ('nav_admin','nav_admin_"&new_page&"','"&new_name&"')")
  call set_session("result",1,"")
  echo 1
end function

function del_user_page()
  dim del_page:del_page = post("del_page","")
  dim del_type:del_type = post("del_type","")
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name = 'new_page' and left(con_value,len('"&del_page&"')+3) = '"&del_page&"{v}'")
  if db.get_record_count() > 0 then
	'删除文件
	set fso = createobject("scripting.filesystemobject")
	fso.deletefile server.mappath(S_ROOT&"index/"&del_page&".asp")
	fso.deletefile server.mappath(S_TPL_PATH&del_page&".asp")
	if del_type = "product" or del_type = "article" then
	  fso.deletefile server.mappath(S_TPL_PATH&"module/"&del_page&"_tree.asp")
	elseif del_type = "about" then
	  fso.deletefile server.mappath(S_TPL_PATH&"module/"&del_page&"_menu.asp")
	end if
	fso.deletefile server.mappath(S_TPL_PATH&"module/"&del_page&"_main.asp")
	fso.deletefolder server.mappath(S_ROOT&"admin/"&del_page)
	set fso = nothing
	'删除代码
	call del_code(S_TPL_PATH&"sinty.xml","<"&del_page,"</"&del_page&">")
	if del_type = "product" or del_type = "article" then
	  call del_code(S_ROOT&"index/module.func.asp","function module_"&del_page&"_tree()","end function")
	elseif del_type = "about" then
	  call del_code(S_ROOT&"index/module.func.asp","function module_"&del_page&"_menu()","end function")
	end if
	call del_code(S_ROOT&"index/module.func.asp","function module_"&del_page&"_main()","end function")
	'删除数据库记录
	call db.reset()
	db.sql_query("delete from "&S_DB_PREFIX&"article where art_type = '"&del_page&"'")
	db.sql_query("delete from "&S_DB_PREFIX&"category where cat_type = '"&del_page&"'")
	db.sql_query("delete from "&S_DB_PREFIX&"config where con_name = 'new_page' and left(con_value,len('"&del_page&"')+3) = '"&del_page&"{v}'")
	db.sql_query("delete from "&S_DB_PREFIX&"config where con_name = 'title' and left(con_value,len('"&del_page&"')+3) = '"&del_page&"{v}'")
	db.sql_query("delete from "&S_DB_PREFIX&"config where con_type = 'nav_stage' and con_name = 'nav_"&del_page&"'")
	db.sql_query("delete from "&S_DB_PREFIX&"config where con_type = 'nav_admin' and con_name = 'nav_admin_"&del_page&"'")
	db.sql_query("delete from "&S_DB_PREFIX&"menu where men_type = '"&del_page&"'")
	db.sql_query("delete from "&S_DB_PREFIX&"menu where men_type = 'admin' and men_url = 'main.asp?dir="&del_page&"&fil=sheet'")
	db.sql_query("delete from "&S_DB_PREFIX&"menu where men_type = 'admin_"&del_page&"' and men_url = 'main.asp?dir="&del_page&"&fil=sheet'")
	db.sql_query("delete from "&S_DB_PREFIX&"menu where men_type = 'admin_"&del_page&"' and men_url = 'main.asp?dir="&del_page&"&fil=add'")
	db.sql_query("delete from "&S_DB_PREFIX&"menu where men_type = 'admin_"&del_page&"' and men_url = 'main.asp?dir="&del_page&"&fil=cat'")
	db.sql_query("delete from "&S_DB_PREFIX&"menu where men_type = 'header' and men_url = '?"&del_page&".html'")
	db.sql_query("delete from "&S_DB_PREFIX&"modules where mod_type = '"&del_page&"'")
	db.sql_query("delete from "&S_DB_PREFIX&"modules where mod_type = '"&del_page&"_sheet'")
	db.sql_query("delete from "&S_DB_PREFIX&"modules where mod_type = '"&del_page&"_form'")
	call set_session("result",1,"")
	echo 1
  end if
end function

'新秀
%>