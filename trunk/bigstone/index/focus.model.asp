<%
class focus_model
  private focus()
  private focus_count
  
  private sub class_initialize()
	
  end sub

  public function get_focus(val)
	dim i,record_count
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"picture where pic_type='focus' and pic_show = 1 and pic_site = left('"+val+"',len(pic_site)) order by pic_top desc,pic_index desc")
	if db.get_record_count() > 0 then
	  record_count = db.get_record_count()
	  focus_count = record_count
	  redim preserve focus(record_count - 1,3)
	else
	  call db.reset()
	  call db.sql_query("select * from "&S_DB_PREFIX&"picture where pic_type='focus' and pic_show = 1 and pic_site = 'default' order by pic_top desc,pic_index desc")
	  if db.get_record_count() > 0 then
		record_count = db.get_record_count()
		focus_count = record_count
		redim preserve focus(record_count - 1,3)
	  end if
	end if
	for i = 0 to record_count - 1
	  focus(i,0) = db.sql_result(i,"pic_path")
	  focus(i,1) = db.sql_result(i,"pic_url")
	  focus(i,2) = db.sql_result(i,"pic_title")
	  focus(i,3) = db.sql_result(i,"pic_site")
	  if left(focus(i,0),5) <> "http:" then focus(i,0) = S_ROOT & focus(i,0)
	next
	get_focus = focus
  end function
  
  public function get_focus_count()
	get_focus_count = focus_count
  end function
  
  public function get_focus_name()
	dim focus_name(2)
	focus_name(0) = "path"
	focus_name(1) = "url"
	focus_name(2) = "title"
	get_focus_name = focus_name
  end function
  
end class
'新秀
%>