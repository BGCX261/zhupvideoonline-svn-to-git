<%
dim rsf
sql="select * from baseSet"
Set rsf= Server.CreateObject("ADODB.Recordset")
rsf.open sql,conn,1,1
%>
<div class="footer">
    <div class="footState">
       <p><span>版权声明：</span><%=rsf("announces")%>
       </p>
      <p> &copy; 2002-2013 Cgfans Studio. Designed by <a href="mailto:cucuang@163.com?subject=从色艺坊(图片)网站发来的信" >老猪</a></p>
    </div> 
</div>
<%                                                     
rsf.close
set rsf=nothing
conn.close
set conn=nothing
%>
</body>
</html>


