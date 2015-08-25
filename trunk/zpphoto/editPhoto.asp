<!--#include file="security.asp"-->
<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<body>
<div class="admCont"> 
  <div class="admTop"></div>
  <div class="admLeft">
      <!--#include file="left.asp"-->
  </div>
  <div class="admRight">
<form method="POST" action="saveEditPhoto.asp?id=<%=request("id")%>" name="cgfans">
    <table class="iptTable">
        <tr> 
            <td colspan="4"  class="doTitle">修改作品</td>
        </tr>
            <%
dim sql
dim rs
 sql="select * from albums where articleid="&request("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
                %>
            <tr> 
              <td width="15%" class="iptText">作品标题</td>
              <td colspan="3"> <input type="text" name="title" size="50" maxlength="100" value="<%=rs("title")%>">
                *</td>
            </tr>
			<tr> 
              <td width="15%" class="iptText"> 作者</td>
              <td width="35%"> <input type="text" name="author" size="16"  maxlength="24" value="<%=rs("author")%>">              </td>
              <td width="15%" class="iptText">来源</td>
              <td width="35%"> <input type="text" name="source" size="24"  maxlength="36" value="<%=rs("source")%>">              </td>
            </tr>
            <tr> 
              <td class="iptText">分类</td>
              <td> <select class="smallSel" name="typeid" size="1">
                  <%               
dim rs1,sql1,sel      
 set rs1=server.createobject("adodb.recordset")      
  sql1="select * from type"      
 rs1.open sql1,conn,1,1      
			  do while not rs1.eof      
                                sel="selected"          
		             response.write "<option  value='"+CStr(rs1("typeID"))+"' name=typeid "      
		           ' Response.Write rs1("typeid")& rs("typeid")      
		             if rs("typeid")=rs1("typeid") then Response.Write sel      
		             Response.Write ">"+rs1("type")+"</option>"+chr(13)+chr(10)      
		             rs1.movenext      
    		          loop      
			rs1.close      
			        %>
                </select></td>
              <td class="iptText">推荐指数</td>
              <td>
              	<script type="text/javascript">
        var conty="";
        var beg=2003;
        var yy=Number(new Date().getFullYear());
        for(i=0;i<=yy-beg;i++){
        var temp=beg+i
        conty+="<option>"+temp+"</option>"
        }
		var gi= "<%=rs("orgTime")%>" 
	
        conty+="<option selected='selected' value="+parseInt(gi)+">"+parseInt(gi)+"</option>"
		
    </script>
    <select name="QYear" size="1" class="Select"> 
        <script type="text/javascript"> document.write(conty)</script>
    </select>年
    <script type="text/javascript">
        var contm="";
		var kk=Number(new Date().getMonth())+1
        for(qq=1;qq<=12;qq++){
        contm+="<option>"+qq+"</option>"
        }
		var giv = "<%=rs("orgTime")%>" 
		ys = giv.indexOf("年")
		ym =  giv.indexOf("月")
		
        contm+="<option selected='selected' value="+kk+">"+giv.substring(ys+1,ym)+"</option>"
    </script>
    <select name="QMonth" size="1">
       <script type="text/javascript"> document.write(contm)</script>
    </select>月 
              </td>
            </tr>
            <tr> 
              <td class="iptText">首页推荐</td>
              <td> <%if rs("best")=true then%>
                是： 
                <input type="radio" name="best" value="1" checked> &nbsp;&nbsp; 
                否： <input type="radio" name="best" value="0">
                
                <%else%>
                是： 
                <input type="radio" name="best" value="1"> &nbsp;&nbsp;否： <input type="radio" name="best" value="0" checked>
                
                <%end if%> </td>
				<td class="iptText">作品类型</td>
			  <td><select name="lx">
			    <option <%if rs("lx")="0" then%> selected <%end if%> value="0"> 图片</option>
			    <option <%if rs("lx")="1" then%> selected <%end if%> value="1"> swf</option>
			    <option <%if rs("lx")="2" then%> selected <%end if%> value="2"> flv</option>
		      </select></td>
            </tr>
            <tr> 
              <td class="iptText">首页图片</td>
              <td colspan="3">
              <input type="text" name="minPic" id="minPic" size="30" maxlength="100" value="<%=rs("bestpic")%>">
                <br><iframe name="ad" frameborder="0" width="100%" height="26" scrolling="auto" src="uploada.asp"></iframe>
                <br><a href="#" onClick="defPic();"><img src="images/defBestPic.jpg" alt="点击可使用默认缩略图"  height="100" id="tou" style="margin:10px;" ></a>
              </td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片1</td>
              <td><input type="text" name="images1" size="30" maxlength="60" value="<%=rs("images1")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片2</td>
              <td><input type="text" name="images2" size="30" maxlength="60" value="<%=rs("images2")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片3</td>
              <td><input type="text" name="images3" size="30" maxlength="60" value="<%=rs("images3")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片4</td>
              <td><input type="text" name="images4" size="30" maxlength="60" value="<%=rs("images4")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片5</td>
              <td><input type="text" name="images5" size="30" maxlength="60" value="<%=rs("images5")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片6</td>
              <td><input type="text" name="images6" size="30" maxlength="60" value="<%=rs("images6")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片7</td>
              <td><input type="text" name="images7" size="30" maxlength="60" value="<%=rs("images7")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片8</td>
              <td><input type="text" name="images8" size="30" maxlength="60" value="<%=rs("images8")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片9</td>
              <td><input type="text" name="images9" size="30" maxlength="60" value="<%=rs("images9")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片10</td>
              <td><input type="text" name="images10" size="30" maxlength="60" value="<%=rs("images10")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片11</td>
              <td><input type="text" name="images11" size="30" maxlength="60" value="<%=rs("images11")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片12</td>
              <td><input type="text" name="images12" size="30" maxlength="60" value="<%=rs("images12")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片13</td>
              <td><input type="text" name="images13" size="30" maxlength="60" value="<%=rs("images13")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片14</td>
              <td><input type="text" name="images14" size="30" maxlength="60" value="<%=rs("images14")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片15</td>
              <td><input type="text" name="images15" size="30" maxlength="60" value="<%=rs("images15")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片16</td>
              <td><input type="text" name="images16" size="30" maxlength="60" value="<%=rs("images16")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片17</td>
              <td><input type="text" name="images17" size="30" maxlength="60" value="<%=rs("images17")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片18</td>
              <td><input type="text" name="images18" size="30" maxlength="60" value="<%=rs("images18")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr class="ipt ipttable">
              <td class="iptText">图片19</td>
              <td><input type="text" name="images19" size="30" maxlength="60" value="<%=rs("images19")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片20</td>
              <td><input type="text" name="images20" size="30" maxlength="60" value="<%=rs("images20")%>" onFocus="AXImg(this)"  /></td>
            </tr>
            
            <tr> 
              <td class="iptText"> 作品说明</td>
              <td colspan="3" > <textarea rows="5" name="content" cols="80"><%if rs("html")=false then
                content=replace(rs("content"),"<br>",chr(13))
                content=replace(content,"&nbsp;"," ")
            else 
                content=rs("content")
            end if        
            response.write content%></textarea>
                <br /><input type="checkbox" name="html" value="on" checked>
                是否支持html              </td>
            </tr>
            <tr> 
              <td colspan="4" class="iptSmt"> <input type="submit" value=" 修 改 " name="cmdok"> 
          <input type="reset" value=" 取 消 " name="cmdcancel">
        <input type="button" value=" 返 回 " name="cmdReturn" onClick="window.location='managePhoto.asp'"></td>
            </tr>
          </table>
	  <div id="preview" style="width:700px; height:450px; z-index:1; overflow: scroll;"> 
            <img id="mages1" /> 
            <img id="mages2" /> 
            <img id="mages3" /> 
            <img id="mages4" /> 
            <img id="mages5" /> 
            <img id="mages6" /> 
            <img id="mages7" /> 
            <img id="mages8" /> 
            <img id="mages9" /> 
            <img id="mages10" /> 
            <img id="mages11" /> 
            <img id="mages12" /> 
            <img id="mages13" /> 
            <img id="mages14" /> 
            <img id="mages15" /> 
            <img id="mages16" /> 
            <img id="mages17" /> 
            <img id="mages18" /> 
            <img id="mages19" /> 
            <img id="mages20" /> 
          </div>
  
</form></div>
</div>
</body>
</html>
<% 
   
 rs.close 
set rs=nothing 
  conn.close  
  set conn=nothing  
%>

<script language="JavaScript">
document.getElementById("tou").src=document.getElementById("minPic").value
<!--
function AXImg(ximg){
	var newpaddr = ximg.name
	var newstr=newpaddr.replace("i","");  
	document.getElementById(newstr).src=ximg.value
	if(document.getElementById(newstr).width > 600){
		document.getElementById(newstr).width = 600
	}
} 



function defPic(){
	document.getElementById("minPic").value="images/defBestPic.jpg"
	document.getElementById("tou").src="images/defBestPic.jpg"
	}
//-->
</script>


