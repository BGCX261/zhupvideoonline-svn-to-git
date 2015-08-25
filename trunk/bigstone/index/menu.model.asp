<%
class menu_model
  private menu()
  private menu_count
  
  private sub class_initialize()
	
  end sub
	
  public function get_menu(men_type)
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"menu where men_type = '"&men_type&"' and men_show = 1 order by men_top desc,men_index desc,men_id asc")
	if db.get_record_count() > 0 then
	  dim record_count:record_count = db.get_record_count()
	  menu_count = record_count
	  redim preserve menu(record_count - 1,2)
	  dim i
	  for i = 0 to record_count - 1
		menu(i,0) = db.sql_result(i,"men_name")
		menu(i,1) = db.sql_result(i,"men_url")
		if left(menu(i,1),5) <> "http:" then
		  menu(i,1) = S_ROOT & menu(i,1)
		  menu(i,2) = 0
		else
		  menu(i,2) = 1
		end if
	  next
	end if
	get_menu = menu
  end function
  
  public function get_menu_count()
	get_menu_count = menu_count
  end function
  public function get_menu_name()
	dim menu_name(2)
	menu_name(0) = "word"
	menu_name(1) = "url"	
	menu_name(2) = "target"
	get_menu_name = menu_name
  end function
	
end class
'新秀
%>