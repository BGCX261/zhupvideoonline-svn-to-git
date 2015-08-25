<%
if get_session("login","") <> session.sessionid then
  dim http_referer:http_referer = cstr(request.servervariables("http_referer"))
  dim server_name:server_name = cstr(request.servervariables("server_name"))
  if mid(http_referer,8,len(server_name)) <> server_name then
	response.redirect(S_ROOT)
	response.end
  end if
  dim user:user = post("user","")
  dim password:password = post("password","")
  if len(user) > 30 then
	user = left(user,30)
  end if
  if len(password) > 30 then
	password = left(password,30)
  end if
  if user = "" or password = "" then
	alert = "请重新登录"
	href = "../index.asp"
	window = "top."
  else
	password = md5(password)
	call db.reset()
	call db.sql_query("select * from admin where adm_user = '"&user&"' and adm_password = '"&password&"'")
	if db.get_record_count() > 0 then
	  call set_session("login",session.sessionid,"")
	  call set_session("user",user,"")
	  call set_session("password",password,"")
	else
	  alert = "用户名不存在或密码不正确"
	  href = "../index.asp"
	  window = "top."
	end if
  end if
else
  user = get_session("user","")
  password = get_session("password","")
  call db.reset()
  call db.sql_query("select * from admin where adm_user = '"&user&"' and adm_password = '"&password&"'")
  if db.get_record_count() = 0 then
	alert = "请重新登录"
	href = "../index.asp"
	window = "top."
  end if
end if

if gets("out","") = "1" then
  unset_session("login")
  unset_session("user")
  unset_session("password")
  alert = "您已经退出系统"
  href = "../index.asp"
end if

'新秀
%>
<!-- #include file="alert.asp" -->