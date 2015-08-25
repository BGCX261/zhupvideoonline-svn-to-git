<%
class link_model
  private img_link()
  private img_count
  private word_link()
  private word_count
  
  private sub class_initialize()
	img_count = 0
	word_count = 0
  end sub

  public function get_img_link()
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"link where lin_img <> 'none' order by lin_index desc")
	if db.get_record_count() > 0 then
	  record_count = db.get_record_count()
	  img_count = record_count
	  redim preserve img_link(record_count - 1,2)
	  for i = 0 to record_count - 1
		img_link(i,0) = db.sql_result(i,"lin_url")
		img_link(i,1) = db.sql_result(i,"lin_img")
		img_link(i,2) = db.sql_result(i,"lin_title")
	  next
	end if
	get_img_link = img_link
  end function
  
  public function get_img_count()
	get_img_count = img_count
  end function
  
  public function get_img_link_name()
	dim img_link_name(2)
	img_link_name(0) = "url"
	img_link_name(1) = "img"
	img_link_name(2) = "title"
	get_img_link_name = img_link_name
  end function
  
  public function get_word_link()
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"link where lin_img = 'none' order by lin_index desc")
	if db.get_record_count() > 0 then
	  dim record_count:record_count = db.get_record_count()
	  word_count = record_count
	  redim preserve word_link(record_count - 1,2)
	  dim i
	  for i = 0 to record_count - 1
		word_link(i,0) = db.sql_result(i,"lin_url")
		word_link(i,1) = db.sql_result(i,"lin_word")
		word_link(i,2) = db.sql_result(i,"lin_title")
	  next
	end if
	get_word_link = word_link
  end function
  
  public function get_word_count()
	get_word_count = word_count
  end function
  
  public function get_word_link_name()
	dim word_link_name(2)
	word_link_name(0) = "url"
	word_link_name(1) = "word"
	word_link_name(2) = "title"
	get_word_link_name = word_link_name
  end function
end class
'新秀
%>