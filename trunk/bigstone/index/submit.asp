<!-- #include file="../include/common.asp" -->
<!-- #include file="../include/sentmail.asp" -->
<!-- #include file="sinty/sinty.asp" -->
<!-- #include file="module.func.asp" -->

<%
dim sinty:set sinty = new sinty_
dim info:info = initial("alert.asp")
dim url:url = get_url()
dim cat,page,id,plu
call set_parameter(0,0,0)
call set_head("alert."&S_SUFFIX)
call select_func()
call sinty.display()
call db.close_db()
set db = nothing
set sinty = nothing

function select_func()
  select case post("cmd","")
	case "leave_word":call leave_word()
	case "research":call research()
  end select
end function

function leave_word()
  dim http_referer:http_referer = cstr(request.servervariables("http_referer"))
  dim server_name:server_name = cstr(request.servervariables("server_name"))
  if mid(http_referer,8,len(server_name)) <> server_name then
	response.redirect(S_ROOT)
	response.end
  end if
  dim user:user = post("user","")
  dim show:show = post("show","")
  dim text:text = post("text","")
  if user = "" or text = "" then
	call sinty.assign("alert_type","empty")
	call sinty.assign("href","?message."&S_SUFFIX)
  else
	if len(user) > 30 then user = left(user,30)
	if show <> "2" then show = "0"
	call db.reset()
	dim values:values = "'"&user&"','"&now()&"','"&text&"','"&show&"'"
	call db.sql_query("insert into message (mes_user,mes_time,mes_text,mes_show) values ("&values&")")
	on error resume next
	if get_config("sentmail") = "1" then
	  dim sit_name:sit_name = get_config("sit_name")
	  dim smtp:smtp = get_config("sentmail_smtp")
	  dim send:send = get_config("sentmail_send")
	  dim password:password = get_config("sentmail_password")
	  dim receive:receive = get_config("sentmail_receive")
	  dim title:title = user & "在" & sit_name & "网站留言"
	  dim main:set mail = new sentmail
	  call mail.init_mail(smtp,send,password,receive,title,text,"")
	  call mail.send()
	end if
	if err then err.clear
	call sinty.assign("alert_type","leave_word")
	call sinty.assign("href","?message."&S_SUFFIX)
  end if
end function

function research()
  dim http_referer:http_referer = cstr(request.servervariables("http_referer"))
  dim server_name:server_name = cstr(request.servervariables("server_name"))
  if mid(http_referer,8,len(server_name)) <> server_name then
	response.redirect(S_ROOT)
	response.end
  end if
  dim arr
  dim k:k = 0
  dim text:text = ""
  dim res_obj:set res_obj = new research_model
  dim res:res = res_obj.get_research()
  if res_obj.get_research_count() > 0 then
	redim preserve question(ubound(res))
	for i = 0 to ubound(res)
	  arr = split(res(i,1),"{v}")
	  text = text & arr(0)
	  for j = 2 to ubound(arr)
		if arr(1) <> "radio" then
		  text = text & "[" & post("res_"&k,"") & "]"
		  k = k + 1
		end if
	  next
	  if arr(1) = "radio" then
		text = text & "[" & post("res_"&k,"") & "]"
		k = k + 1
	  end if
	  text = text & "{v}"
	next
	call db.reset()
	dim ip:ip = get_ip()
	dim values:values = "'"&ip&"','"&now()&"','"&text&"'"
	call db.sql_query("insert into research (res_ip,res_time,res_text) values ("&values&")")
	on error resume next
	if get_config("sentmail") = "1" then
	  dim sit_name:sit_name = get_config("sit_name")
	  dim smtp:smtp = get_config("sentmail_smtp")
	  dim send:send = get_config("sentmail_send")
	  dim password:password = get_config("sentmail_password")
	  dim receive:receive = get_config("sentmail_receive")
	  dim title:title = "IP为"& ip &"的访客在" & sit_name & "网站提交了调查问卷"
	  text = replace(replace(text,"[]",""),"{v}","<br>")
	  dim main:set mail = new sentmail
	  call mail.init_mail(smtp,send,password,receive,title,text,"")
	  call mail.send()
	end if
	if err then err.clear
  end if
  call sinty.assign("alert_type","research")
  call sinty.assign("href","?message."&S_SUFFIX)
end function
'新秀
%>