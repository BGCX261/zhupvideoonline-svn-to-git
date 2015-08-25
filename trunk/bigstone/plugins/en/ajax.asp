<%
select case request.form("cmd")
  case "sel_cn":call sel_cn()
  case "sel_en":call sel_en()
end select

function sel_cn()
  if session("plugin") <> "" then session.contents.remove("plugin")
end function

function sel_en()
  session("plugin") = "en"
end function

'新秀
%>