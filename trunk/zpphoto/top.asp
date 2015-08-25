    <div class="logo"> <a href="http://www.zhuping.net" target="_top"> <img alt="www.zhuping.net" src="images/zpLogo.gif" /></a> </div>
    <form action="search.asp" method="post" name="search" onSubmit="return checkinput()">
    <div class="menu">
        <div class="menubox">
    <%
        set rst=server.CreateObject("ADODB.RecordSet") 
        rst.open "select * from type",conn,1
        if rst.EOF then
        response.write "没有栏目："
        else
    %>
      <a href="index.asp">图集首页</a>
      <%do while NOT rst.EOF%>
      | <a href="photoType.asp?typeid=<%=rst("typeid")%>"><%=rst("type")%></a>
     <%                                
        rst.MoveNext         
        loop
     %>         
      | <a href="brief.asp">关于</a>
     <% 
		end if            
		rst.close
		set rst=nothing         
     %>
        </div>
        <div class="searchbox">
           <input type="text" name="keyword" size="12" maxlength="30" value="请输入关键字" onBlur="if(this.value==''){this.value='请输入关键字'};"  onfocus="if(this.value=='请输入关键字'){this.value=''};" class="input">
           <input type="hidden" name="typeid" value="&typeid&">
           <input type="submit" onMouseOut="this.className='searchbtnout'" onMouseOver="this.className='searchbtnover'"   name="search" value="搜 索"  class="searchbtnout">
        </div>
    </div>
    </form>

