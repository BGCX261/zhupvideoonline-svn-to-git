<!--#include file="security.asp"-->
<!--#include file="dbConn.asp"-->
<%
dim a
a=request.form
response.write a
%>
<!--#include file="head.asp"-->
<body>
<div class="admCont"> 
  <div class="admTop"></div> 
  <div class="admLeft">
      <!--#include file="left.asp"-->
  </div>
  <div class="admRight">
<form method="POST" action="savePhoto.asp" name="cgfans">
     <table class="iptTable">
            <tr> 
              <td colspan="4" class="doTitle">添加作品</td>
            </tr>
            <tr> 
              <td colspan="4" class="alm">若要上传本地图片，请先<a href="upload/upFile.asp">上传图片</a>，图片地址会自动回填到图片地址框中。 若是网上的图片，请把图片地址按顺序填入下面的框中</td>
            </tr>
            <tr> 
              <td class="iptText">作品标题</td>
              <td colspan="3"> <input type="text" name="title" size="50"  maxlength="100"><span class="chkAlm"></span></td>
            </tr>
			<tr> 
              <td width="15%" class="iptText"> 作者</td>
              <td width="35%"> <input type="text" name="author" size="16"  maxlength="24" value="方言">              </td>
              <td width="15%" class="iptText">来源</td>
              <td width="35%"> <input type="text" name="source" size="24"  maxlength="36" value="本站">              </td>
            </tr>
            <tr> 
              <td class="iptText">分类</td>
              <td> <select name="typeid" size="1">
                  <%                
dim rs,sql,sel           
 set rs=server.createobject("adodb.recordset")           
  sql="select * from type"           
 rs.open sql,conn,1,1           
			  do while not rs.eof           
                                sel="selected"               
		             response.write "<option " & sel & " value='"+CStr(rs("typeID"))+"' name=typeid>"+rs("type")+"</option>"+chr(13)+chr(10)           
		             rs.movenext           
    		          loop           
			rs.close
			set rs=nothing           
			      %>
                  <option selected value="0" name="typeid">请选择图片分类</option>
                </select> &nbsp; </td>
              <td class="iptText">拍摄时期</td>
              <td>
	<script type="text/javascript">
        var conty="";
        var beg=2003;
        var yy=Number(new Date().getFullYear());
        for(i=0;i<=yy-beg;i++){
        var temp=beg+i
        conty+="<option>"+temp+"</option>"
        }
        conty+="<option selected='selected' value="+yy+">"+yy+"</option>"
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
        contm+="<option selected='selected' value="+kk+">"+kk+"</option>"
    </script>
    <select name="QMonth" size="1">
       <script type="text/javascript"> document.write(contm)</script>
    </select>月 
    
              </td>
            </tr>
            <tr> 
              <td class="iptText">首页推荐</td>
              <td > <input type="radio" value="true" name="best">
                是 
                <input type="radio" value="false" checked name="best">
                否</td>
				<td class="iptText">作品类型</td>
			  <td><select name="lx">
			    <option value='0' selected>图片</option>
			    <option value='1'>swf</option>
			    <option value='2'>flv</option>
		      </select></td>
            </tr>
            <tr>
              <td class="iptText">首页缩略图</td>
              <td height="160" colspan="3"><input type="text" name="minPic" id="minPic" size="30" maxlength="100" value="">
                <br><iframe name="ad" frameborder="0" width="100%" height="26" scrolling="auto" src="uploada.asp"></iframe>
                <br><a href="#" onClick="defPic();"><img src="images/defBestPic.jpg" alt="点击可使用默认缩略图" width="80" height="120" id="tou" style="margin:10px;" ></a></td>
            </tr>
            <tr>
              <td class="iptText">图片1</td>
              <td><input type="text" name="images1" size="30" maxlength="60" value="<%=request("tabs1")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片2</td>
              <td><input type="text" name="images2" size="30" maxlength="60" value="<%=request("tabs2")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片3</td>
              <td><input type="text" name="images3" size="30" maxlength="60" value="<%=request("tabs3")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片4</td>
              <td><input type="text" name="images4" size="30" maxlength="60" value="<%=request("tabs4")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片5</td>
              <td><input type="text" name="images5" size="30" maxlength="60" value="<%=request("tabs5")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片6</td>
              <td><input type="text" name="images6" size="30" maxlength="60" value="<%=request("tabs6")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片7</td>
              <td><input type="text" name="images7" size="30" maxlength="60" value="<%=request("tabs7")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片8</td>
              <td><input type="text" name="images8" size="30" maxlength="60" value="<%=request("tabs8")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片9</td>
              <td><input type="text" name="images9" size="30" maxlength="60" value="<%=request("tabs9")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片10</td>
              <td><input type="text" name="images10" size="30" maxlength="60" value="<%=request("tabs10")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片11</td>
              <td><input type="text" name="images11" size="30" maxlength="60" value="<%=request("tabs11")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片12</td>
              <td><input type="text" name="images12" size="30" maxlength="60" value="<%=request("tabs12")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片13</td>
              <td><input type="text" name="images13" size="30" maxlength="60" value="<%=request("tabs13")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片14</td>
              <td><input type="text" name="images14" size="30" maxlength="60" value="<%=request("tabs14")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片15</td>
              <td><input type="text" name="images15" size="30" maxlength="60" value="<%=request("tabs15")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片16</td>
              <td><input type="text" name="images16" size="30" maxlength="60" value="<%=request("tabs16")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片17</td>
              <td><input type="text" name="images17" size="30" maxlength="60" value="<%=request("tabs17")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片18</td>
              <td><input type="text" name="images18" size="30" maxlength="60" value="<%=request("tabs18")%>" onFocus="AXImg(this)" /></td>
            </tr>
            <tr>
              <td class="iptText">图片19</td>
              <td><input type="text" name="images19" size="30" maxlength="60" value="<%=request("tabs19")%>" onFocus="AXImg(this)" /></td>
              <td class="iptText">图片20</td>
              <td><input type="text" name="images20" size="30" maxlength="60" value="<%=request("tabs20")%>" onFocus="AXImg(this)" /></td>
            </tr>
            
            <tr> 
              <td class="iptText">作品说明</td>
              <td colspan="3"> <textarea rows="5" name="content" cols="80">暂无</textarea>
                <br /><input type="checkbox" name="html" value="on">
                是否支持html              </td>
            </tr>
            <tr> 
              <td colspan="4" class="iptSmt"><input type="submit" value=" 添 加 " name="cmdok"> 
          <input type="reset" value=" 清 除 " name="cmdcancel"></td>
            </tr>
          </table>
          <div id="preview" style=" width:700px; height:450px; z-index:1; overflow: scroll;"> 
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
</form>
</div>
</div>
</body>
</html>
<script language="JavaScript">
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