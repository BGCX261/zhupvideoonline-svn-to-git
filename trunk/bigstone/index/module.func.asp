<!-- #include file="menu.model.asp" -->
<!-- #include file="focus.model.asp" -->
<!-- #include file="article.model.asp" -->
<!-- #include file="contact.model.asp" -->
<!-- #include file="link.model.asp" -->
<!-- #include file="footer.model.asp" -->
<!-- #include file="tree.model.asp" -->
<!-- #include file="message.model.asp" -->
<!-- #include file="search.model.asp" -->
<!-- #include file="research.model.asp" -->

<%
function initial(byval fil)
  call sinty.set_include("../../../include/config.asp")
  call sinty.set_include("../../../include/function.asp")
  call sinty.set_include("../../../languages/"&S_LANG&".asp")
  call sinty.set_sinty_path(S_ROOT&"index/sinty/")
  call sinty.set_rewrite(S_REWRITE)
  initial = sinty.get_tpl_info(S_TPL_PATH&fil)
end function
function set_head(byval tab)
  if tab = "recruit" then tab = "article"
  dim val:val = left(tab,3)
  if url = "" then url = "index." & S_SUFFIX
  dim arr:arr = split(mid(url,1,instr(url,".") - 1),"_")
  call sinty.assign("url_page",arr(0))
  call sinty.assign("site_title",get_config("sit_title"))
  call sinty.assign("channel_title",get_channel_title())  
  call sinty.assign("cat_id",cat)
  call sinty.assign("id",id)
  call sinty.assign("url",url)
  call sinty.assign("suffix",S_SUFFIX)
  if id = 0 then
	call sinty.assign("keywords",get_config("sit_keywords"))
	call sinty.assign("describe",get_config("sit_description"))
  else
	call sinty.assign("page_title",get_data(tab,id,val&"_title"))
	call sinty.assign("keywords",get_data(tab,id,val&"_keywords"))
	call sinty.assign("describe",get_data(tab,id,val&"_description"))
	call sinty.assign("cat_id",get_data(tab,id,val&"_cat_id"))
	call sinty.assign("cat_name",get_data("category",get_data(tab,id,val&"_cat_id"),"cat_name"))
  end if
  if cat <> 0 and id = 0 then call sinty.assign("cat_name",get_data("category",cat,"cat_name"))
end function
function get_channel_title()
  dim i,arr
  dim return:return = ""
  dim priority:priority = 0
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name='title'")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	for i = 0 to record_count - 1
	   arr = split(db.sql_result(i,"con_value"),"{v}")
	   if left(url,len(arr(0))) = arr(0) and int(arr(2)) >= priority then
		 return = arr(1)
		 priority = int(arr(2))
	   end if
	next
  end if
  get_channel_title = return
end function
function found_html(byval fil)
  if S_STATIC > 0 then
	fil = S_ROOT&"html/"&fil
	dim path
	if url = "" then
	  path = server.mappath(fil)
	else
	  path = server.mappath(mid(fil,1,instrrev(fil,"/"))&url)
	end if
	call sinty.found_html(path)
  end if
end function
function set_parameter(a,b,c)
  cat = a:page = b:id = c
  if url <> "" then
	dim str:str = mid(url,1,instr(url,".") - 1)
	dim arr:arr = split(str,"_")
	select case ubound(arr)
	  case 1:id = arr(1)
	  case 2:page = arr(1):id = arr(2)
	  case 3:cat = arr(1):page = arr(2):id = arr(3)
	end select
	dim len_:len_ = instrrev(url,".") - instr(url,".")
	if len_ > 0 then plu = mid(url,instr(url,".") + 1,len_ - 1) else plu = ""
  end if
  if cat = "" then cat = 0 else cat = int(cat)
  if page = ""  then page = 0 else page = int(page)
  if id = "" then id = 0 else id = int(id)
end function
function set_link(byval page_sum,byval record_sum)
  call sinty.assign("page_sum",page_sum)
  call sinty.assign("record_sum",record_sum)
  call sinty.assign("cat",cat)
  call sinty.assign("page",num_bound(1,page_sum,page))
end function
function select_module(arr)
  dim i,func
  for i = 0 to ubound(arr)
	func = "module_" & arr(i) & "()"
	eval(func)
  next
end function

function module_header()
  dim obj:set obj = new menu_model
  call sinty.assign("nav",obj.get_menu("header"))
  call sinty.assign("nav_u",obj.get_menu_count() - 1)
  call sinty.assign("nav_name",obj.get_menu_name())
  set obj = nothing
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  set obj = new menu_model
  call sinty.assign("nav_about",obj.get_menu("about"))
  call sinty.assign("nav_about_u",obj.get_menu_count() - 1)
  call sinty.assign("nav_about_name",obj.get_menu_name())
  set obj = nothing
  set obj = new tree_model
  call sinty.assign("nav_pro",obj.get_tree("product"))
  call sinty.assign("nav_pro_u",obj.get_tree_count() - 1)
  call sinty.assign("nav_pro_name",obj.get_tree_name())
  set obj = nothing
  set obj = new tree_model
  call sinty.assign("nav_art",obj.get_tree("article"))
  call sinty.assign("nav_art_u",obj.get_tree_count() - 1)
  call sinty.assign("nav_art_name",obj.get_tree_name())
  set obj = nothing
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
end function
function module_focus()
  dim obj:set obj = new focus_model
  call sinty.assign("focus",obj.get_focus(url))
  call sinty.assign("focus_u",obj.get_focus_count() - 1)
  call sinty.assign("focus_name",obj.get_focus_name())
  set obj = nothing
end function
function module_notice()
  dim id:id = get_id("article","art_type","notice")
  call sinty.assign("notice",get_data("article",id,"art_text"))
end function
function module_contact()
  dim obj:set obj = new contact_model
  call sinty.assign("contact",obj.get_contact())
  call sinty.assign("contact_u",obj.get_contact_count() - 1)
  call sinty.assign("contact_name",obj.get_contact_name())
  set obj = nothing
end function
function module_link()
  dim obj:set obj = new link_model
  call sinty.assign("img_link",obj.get_img_link())
  call sinty.assign("img_link_u",obj.get_img_count() - 1)
  call sinty.assign("img_link_name",obj.get_img_link_name())
  call sinty.assign("word_link",obj.get_word_link())
  call sinty.assign("word_link_u",obj.get_word_count() - 1)
  call sinty.assign("word_link_name",obj.get_word_link_name())
  set obj = nothing
end function
function module_about()
  dim obj:set obj = new article_model
  call obj.set_fields("art_title,art_text")
  call obj.set_art_type("about")
  call obj.set_page_size(1)
  dim arr:arr = obj.get_article()
  dim about_filter:about_filter = get_config("about_filter")
  dim about_len:about_len = get_config("about_len")
  if ubound(arr) = 0 then
    if about_filter = "1" then
	  arr(0,1) = cut_str(remove_html(arr(0,1)),about_len)
	else
	  arr(0,1) = repair_html(cut_str(arr(0,1),about_len))
	end if
	call sinty.assign("about_title",arr(0,0))
	call sinty.assign("about_text",arr(0,1))
  else
	call sinty.assign("about_title","")
	call sinty.assign("about_text","")
  end if
  set obj = nothing
end function
function module_recommend()
  dim obj:set obj = new article_model
  call obj.set_art_type("product")
  call obj.set_fields("art_id,art_title,art_reduce_img")
  call obj.set_top(1)
  call obj.set_page_size(30)
  call sinty.assign("recommend",obj.get_article())
  call sinty.assign("recommend_u",obj.get_page_size() - 1)
  call sinty.assign("recommend_name",obj.get_article_name())
  set obj = nothing
end function
function module_news()
  dim obj:set obj = new article_model
  call obj.set_art_type("article")
  call obj.set_fields("art_id,art_title,art_add_time")
  call obj.set_page_size(get_config("index_list_len"))
  call sinty.assign("news",obj.get_article())
  call sinty.assign("news_u",obj.get_page_size() - 1)
  call sinty.assign("news_name",obj.get_article_name())
  set obj = nothing
end function
function module_product()
  dim obj:set obj = new article_model
  call obj.set_art_type("product")
  call obj.set_fields("art_id,art_title,art_add_time")
  call obj.set_page_size(get_config("index_list_len"))
  call sinty.assign("product",obj.get_article())
  call sinty.assign("product_u",obj.get_page_size() - 1)
  call sinty.assign("product_name",obj.get_article_name())
  set obj = nothing
end function
function module_research()
  dim obj:set obj = new research_model
  call sinty.assign("research",obj.get_research())
  call sinty.assign("research_u",obj.get_research_count() - 1)
  call sinty.assign("research_name",obj.get_research_name())
  set obj = nothing
end function
function module_footer()
  dim obj
  set obj = new menu_model
  call sinty.assign("footer_nav",obj.get_menu("footer"))
  call sinty.assign("footer_nav_u",obj.get_menu_count() - 1)
  call sinty.assign("footer_nav_name",obj.get_menu_name())
  set obj = nothing
  set obj = new footer_model
  call sinty.assign("sit_domain",obj.get_info("sit_domain"))
  call sinty.assign("sit_code",obj.get_info("sit_code"))
  call sinty.assign("sit_code_url",obj.get_info("sit_code_url"))
  call sinty.assign("sit_tec",obj.get_info("sit_tec"))
  call sinty.assign("sit_tec_url",obj.get_info("sit_tec_url"))
  call sinty.assign("sit_count_code",obj.get_info("sit_count_code"))
  set obj = nothing
end function
function module_service()
  dim art_id:art_id = get_id("article","art_type","service")
  call sinty.assign("service",get_data("article",art_id,"art_text"))
end function

function module_about_menu()
  dim obj:set obj = new menu_model
  call sinty.assign("menu",obj.get_menu("about"))
  call sinty.assign("menu_u",obj.get_menu_count() - 1)
  call sinty.assign("menu_name",obj.get_menu_name())
  set obj = nothing
end function
function module_about_main()
  dim obj:set obj = new article_model
  call obj.set_art_type("about")
  call obj.set_fields("art_title,art_text")
  call obj.set_art_id(id)
  call sinty.assign("about",obj.get_article())
  call sinty.assign("about_u",obj.get_page_size() - 1)
  call sinty.assign("about_name",obj.get_article_name())
  set obj = nothing
end function
function module_product_tree()
  dim obj:set obj = new tree_model
  call sinty.assign("tree",obj.get_tree("product"))
  call sinty.assign("tree_u",obj.get_tree_count() - 1)
  call sinty.assign("tree_name",obj.get_tree_name())
  set obj = nothing
end function
function module_product_main()
  dim obj:set obj = new article_model
  call obj.set_art_type("product")
  if id = 0 then
    call sinty.assign("show_sheet",1)
	call obj.set_family(get_cat_family("product",cat))
	call obj.set_fields("art_id,art_title,art_reduce_img")
	call obj.set_cat_id(cat)
	call obj.set_page_size(get_config("img_list_len"))
	call obj.set_page_num(page)
  else
	call sinty.assign("show_sheet",0)
	call obj.set_fields("art_id,art_title,art_img,art_text,att_number,att_price")
	call obj.set_art_id(id)
	call sinty.assign("prev_id",obj.get_prev("id"))
	call sinty.assign("prev_title",obj.get_prev("title"))
	call sinty.assign("next_id",obj.get_next("id"))
	call sinty.assign("next_title",obj.get_next("title"))
  end if
  call sinty.assign("product",obj.get_article())
  call sinty.assign("product_u",obj.get_page_size() - 1)
  call sinty.assign("product_name",obj.get_article_name())
  if id = 0 then call set_link(obj.get_page_sum(),obj.get_record_sum())
  set obj = nothing
end function

private function module_article_tree()
  dim obj:set obj = new tree_model
  call sinty.assign("tree",obj.get_tree("article"))
  call sinty.assign("tree_u",obj.get_tree_count() - 1)
  call sinty.assign("tree_name",obj.get_tree_name())
  set obj = nothing
end function
private function module_article_main()
  dim obj:set obj = new article_model
  call obj.set_art_type("article")
  if id = 0 then
    call sinty.assign("show_sheet",1)
	call obj.set_family(get_cat_family("article",cat))
	call obj.set_fields("art_id,art_title,art_add_time")
	call obj.set_cat_id(cat)
	call obj.set_page_size(get_config("art_list_len"))
	call obj.set_page_num(page)
  else
	call sinty.assign("show_sheet",0)
	call obj.set_fields("art_id,art_title,att_author,art_add_time,art_text")
	call obj.set_art_id(id)
	call sinty.assign("prev_id",obj.get_prev("id"))
	call sinty.assign("prev_title",obj.get_prev("title"))
	call sinty.assign("next_id",obj.get_next("id"))
	call sinty.assign("next_title",obj.get_next("title"))
  end if
  call sinty.assign("article",obj.get_article())
  call sinty.assign("article_u",obj.get_page_size() - 1)
  call sinty.assign("article_name",obj.get_article_name())
  if id = 0 then call set_link(obj.get_page_sum(),obj.get_record_sum())
  set obj = nothing
end function

function module_recruit_main()
  dim obj:set obj = new article_model
  call obj.set_art_type("recruit")
  if id = 0 then
    call sinty.assign("show_sheet",1)
	call obj.set_family(get_cat_family("article",cat))
	call obj.set_fields("art_id,art_title,art_add_time,art_text")
	call obj.set_cat_id(cat)
	call obj.set_page_size(get_config("art_list_len"))
	call obj.set_page_num(page)
  else
	call sinty.assign("show_sheet",0)
	call obj.set_fields("art_id,art_title,att_author,art_add_time,art_text")
	call obj.set_art_id(id)
	call sinty.assign("prev_id",obj.get_prev("id"))
	call sinty.assign("prev_title",obj.get_prev("title"))
	call sinty.assign("next_id",obj.get_next("id"))
	call sinty.assign("next_title",obj.get_next("title"))
  end if
  call sinty.assign("article",obj.get_article())
  call sinty.assign("article_u",obj.get_page_size() - 1)
  call sinty.assign("article_name",obj.get_article_name())
  if id = 0 then call set_link(obj.get_page_sum(),obj.get_record_sum())
  set obj = nothing
end function

function module_download_main()
  dim obj:set obj = new article_model
  call obj.set_fields("art_title,art_text,art_add_time,att_url")
  call obj.set_art_type("download")
  call obj.set_page_size(5)
  call obj.set_page_num(page)
  call sinty.assign("download",obj.get_article())
  call sinty.assign("download_u",obj.get_page_size() - 1)
  call sinty.assign("download_name",obj.get_article_name())
  if id = 0 then call set_link(obj.get_page_sum(),obj.get_record_sum())
  set obj = nothing
end function

function module_message_main()
  dim obj:set obj = new message_model
  call obj.set_fields("mes_user,mes_time,mes_text,mes_reply")
  call obj.set_page_size(5)
  call obj.set_page_num(page)
  call sinty.assign("message",obj.get_message())
  call sinty.assign("message_u",obj.get_page_size() - 1)
  call sinty.assign("message_name",obj.get_message_name())
  if id = 0 then call set_link(obj.get_page_sum(),obj.get_record_sum())
  set obj = nothing
end function

function module_search_main()
  dim obj:set obj = new search_model
  call obj.set_fields("art_id,art_title,art_reduce_img")
  call obj.set_page_size(get_config("img_list_len"))
  call obj.set_page_num(page)
  call sinty.assign("search",obj.get_product(keywords))
  call sinty.assign("search_u",obj.get_page_size() - 1)
  call sinty.assign("search_name",obj.get_search_name())
  if id = 0 then call set_link(obj.get_page_sum(),obj.get_record_sum())
  set obj = nothing
end function

'new function

'新秀
%>