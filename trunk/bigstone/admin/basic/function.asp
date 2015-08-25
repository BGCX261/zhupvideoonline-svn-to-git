<%
function get_contact()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name='contact'")
  dim record_count:record_count = db.get_record_count()
  dim i,arr()
  if record_count > 0 then
	redim preserve arr(record_count - 1,2)
	for i = 0 to record_count - 1
	  dim val:val = split(db.sql_result(i,"con_value"),"{v}")
	  arr(i,0) = val(0)
	  arr(i,1) = val(1)
	  arr(i,2) = db.sql_result(i,"con_id")
	next
  else
	redim preserve arr(-1)
  end if
  get_contact = arr
end function

function get_show()
  dim sinty_xml:sinty_xml = server.mappath(S_TPL_PATH&"sinty.xml")
  dim xml:set xml = server.createobject("microsoft.xmldom")
  xml.load(sinty_xml)
  dim i,arr()
  dim k:k = 0
  dim len_a:len_a = xml.selectsinglenode("site").childnodes.length
  for i = 0 to len_a - 1
    dim node_name:node_name = xml.selectsinglenode("site").childnodes.item(i).nodename
	dim name_a:name_a = xml.selectsinglenode("site").childnodes.item(i).getattribute("name")
	if name_a <> "" then
	  dim len_b:len_b = xml.selectsinglenode("site").childnodes.item(i).childnodes.length
	  dim j
	  for j = 0 to len_b - 1
		dim name_b:name_b = xml.selectsinglenode("site").childnodes.item(i).childnodes.item(j).getattribute("name")
		dim node_text:node_text = xml.selectsinglenode("site").childnodes.item(i).childnodes.item(j).text
		dim checked_1,checked_0
		if left(node_text,1) <> "*" then
		  checked_1 = "checked"
		  checked_0 = ""
		else
		  checked_1 = ""
		  checked_0 = "checked"
		  node_text = replace(node_text,"*","")
		end if
		redim preserve arr(3,k)
		arr(0,k) = name_a & name_b
		arr(1,k) = node_name & "-" & node_text
		arr(2,k) = checked_1
		arr(3,k) = checked_0
		k = k + 1
	  next
	end if
  next
  set xml = nothing
  get_show = arr
end function

'新秀
%>