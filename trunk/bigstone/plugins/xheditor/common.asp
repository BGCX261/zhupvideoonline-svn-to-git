<%
function plu_xheditor(val)
  dim upload_path:upload_path = S_ROOT & "plugins/xheditor/upload.asp"
  %>
  <script type="text/javascript" src="<%=S_ROOT%>plugins/xheditor/jquery-1.4.4.min.js"></script>
  <script type="text/javascript" src="<%=S_ROOT%>plugins/xheditor/xheditor-1.1.12-zh-cn.min.js"></script>
  <textarea id="elm1" name="editor" class="xheditor" rows="12" cols="80"><%=val%></textarea>
  <script language="javascript">
  $('#elm1').xheditor({upLinkUrl:"<%=upload_path%>",upLinkExt:"zip,rar,txt",upImgUrl:"<%=upload_path%>",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"<%=upload_path%>",upFlashExt:"swf",upMediaUrl:"<%=upload_path%>",upMediaExt:"avi"});
  </script>
  <%
end function
'新秀
%>