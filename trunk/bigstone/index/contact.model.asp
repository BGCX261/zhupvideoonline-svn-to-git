<%
class contact_model
  private contact()
  private contact_count
  
  private sub class_initialize()
	
  end sub
	
  public function get_contact()
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name='contact'")
	if db.get_record_count() > 0 then
	  dim record_count:record_count = db.get_record_count()
	  contact_count = record_count
	  redim preserve contact(record_count - 1,1)
	  dim i
	  for i = 0 to record_count - 1
		dim val:val = split(db.sql_result(i,"con_value"),"{v}")
		contact(i,0) = val(0)
		contact(i,1) = val(1)
	  next
	end if
	get_contact = contact
  end function
  
  public function get_contact_count()
	get_contact_count = contact_count
  end function
  public function get_contact_name()
	dim contact_name(1)
	contact_name(0) = "cue"
	contact_name(1) = "content"	
	get_contact_name = contact_name
  end function
	
end class
'新秀
%>