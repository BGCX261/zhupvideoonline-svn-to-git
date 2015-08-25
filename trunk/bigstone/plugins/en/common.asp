<%
function plu_en_reset()
  S_DB_PATH = "plugins/en/"
  S_DB_NAME = "#data_en.mdb"
  S_LANG = "en-us"
  S_SUFFIX = "en.html"
end function

function plu_en_select(val)
  dim radio_cn,radio_en
  if session("plugin") <> "en" then
	radio_cn = "checked"
	radio_en = ""
  else
	radio_cn = ""
	radio_en = "checked"
  end if
  %>
  <div style="position:absolute;right:5px;bottom:3px;color:#FFF;">
	当前数据库：
	中文版<input name="plu_en_radio" type="radio" onClick="plu_en_lang('cn')" <%=radio_cn%> />&nbsp;&nbsp;
	英文版<input name="plu_en_radio" type="radio" onClick="plu_en_lang('en')" <%=radio_en%> />&nbsp;&nbsp;
  </div>
  <script language="javascript">
  function plu_en_lang(val)
  {
	  ajax("post","../../plugins/en/ajax.asp","cmd=sel_"+val,
	  function(data)
	  {
		  //alert(data);
	  });
  }
  </script>
  <%
end function

'新秀
%>