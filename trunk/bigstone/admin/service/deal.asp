<!-- #include file="../system/common.asp" -->
<%
dim func:func = post("cmd","") & "()"
eval(func)

function del_mes()
  dim arr:arr = split(post("id",""),",")
  dim id:id = arr(0)
  dim page
  if arr(1) = "" then page = 1 else page = arr(1)
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"message where mes_id = "&id)
  call set_session("result",1,"")
  echo page
end function

function reply_mes()
  dim id:id = post("id","")
  dim page:page = post("page","")
  dim mes_reply:mes_reply = post("mes_reply","")
  if page = "" then page = 1
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"message set mes_reply = '"&mes_reply&"' where mes_id = "&id)
  call set_session("result",1,"")
  echo page
end function

function verify_mes()
  dim id:id = post("id","")
  dim mes_show:mes_show = post("mes_show","")
  if mes_show <> 0 and mes_show <> 1 then mes_show = 0
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"message set mes_show = '"&mes_show&"' where mes_id = "&id)
  echo 1
end function

function edit_question()
  dim i,question,input_type,answer,id,val
  for i = 1 to request.form("question").count
	question = strict(request.form("question")(i))
	input_type = strict(request.form("input_type")(i))
	answer = strict(request.form("answer")(i))
	id = strict(request.form("id")(i))
	if question <> "" and input_type <> "" and answer <> "" then
	  val = question&"{v}"&input_type&"{v}"&replace(answer,"|","{v}")
	  call db.reset()
	  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_id="&id)
	end if
  next
  call set_session("result",1,"")
end function

function add_question()
  dim question,input_type,answer,val
  question = post("question","")
  input_type = post("input_type","")
  answer = post("answer","")
  if question <> "" and input_type <> "" and answer <> "" then
	val = question&"{v}"&input_type&"{v}"&replace(answer,"|","{v}")
	call db.reset()
	call db.sql_query("insert into "&S_DB_PREFIX&"config (con_name,con_value) values ('research','"&val&"')")
  end if
  call set_session("result",1,"")
  echo 1
end function

function del_question()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"config where con_id = "&id)
  call set_session("result",1,"")
  echo 1
end function

function del_research()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"research where res_id = "&id)
  call set_session("result",1,"")
  echo 1
end function

function edit_notice()
  dim art_id:art_id = post("id","")
  dim art_text:art_text = post("editor","loose")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"article set art_title='"&art_title&"',art_text='"&art_text&"' where art_id = "&art_id)
  call set_session("result",1,"")
  echo 1
end function

function edit_online()
  dim art_id:art_id = post("id","")
  dim art_text:art_text = post("editor","loose")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"article set art_text='"&art_text&"' where art_id = "&art_id)
  call set_session("result",1,"")
  echo 1
end function

function del_link()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"link where lin_id = "&id)
  call set_session("result",1,"")
  echo 1
end function

function add_link()
  dim lin_word:lin_word = post("lin_word","")
  dim lin_url:lin_url = post("lin_url","")
  dim lin_img:lin_img = post("lin_img","")
  dim lin_index:lin_index = post("lin_index","")
  dim lin_title:lin_title = post("lin_title","")
  if replace(lin_img,"http://","") = "" then lin_img = "none"
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"link (lin_word,lin_url,lin_img,lin_index,lin_title) values ('"&lin_word&"','"&lin_url&"','"&lin_img&"','"&lin_index&"','"&lin_title&"')")
  call set_session("result",1,"")
  echo 1
end function

function edit_link()
  dim lin_id:lin_id = post("lin_id","")
  dim lin_word:lin_word = post("lin_word","")
  dim lin_url:lin_url = post("lin_url","")
  dim lin_img:lin_img = post("lin_img","")
  dim lin_index:lin_index = post("lin_index","")
  dim lin_title:lin_title = post("lin_title","")
  if replace(lin_img,"http://","") = "" then lin_img = "none"
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"link set lin_word='"&lin_word&"',lin_url='"&lin_url&"',lin_img='"&lin_img&"',lin_index='"&lin_index&"',lin_title='"&lin_title&"' where lin_id = "&lin_id)
  call set_session("result",1,"")
  echo 1
end function

call db.close_db()
set db = nothing
'新秀
%>