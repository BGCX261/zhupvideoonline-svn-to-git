<!--#include file="security.asp"-->
<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<script type="text/javascript" src="ubbeditor/ubbEditor.js"></script>
<body>
<div class="admCont"> 
  <div class="admTop"></div>
  <div class="admLeft">
      <!--#include file="left.asp"-->
  </div>
  <div class="admRight">
  <div>简介设置</div>
 <form method="POST" action="saveBriefInfo.asp" >

        <%set rs=server.createobject("adodb.recordset")
  rs.open "select * from baseSet",conn,1,3
  %>
          <textarea name="briefInfo" id="briefInfo" style="WIDTH: 680px; HEIGHT: 220px"><%=rs("briefInfo")%></textarea>
				<script type="text/javascript">
				var nEditor = new ubbEditor('briefInfo');
				nEditor.tEditUBBMode = 0;//是否启用html编辑器，0=启用。
				nEditor.tLang = 'zh-cn';
				nEditor.tInit('nEditor', 'ubbeditor/');
				nEditor.tGetHTML();
                </script>
          <input type="submit" value=" 修 改 " name="briefOk">
              <input type="reset" value=" 重 写 " name="cmdcancel">
      </form>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
    </div>
</div>
</body>
</html>


