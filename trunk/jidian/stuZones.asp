<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/Top.asp" -->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_stuZone.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
      <div class="mbox">
        <div class="artlist"><div class="positiontitle"><img src="img/title_ico.gif" /> 师生天地</div>
<% 
'=================================================
'过程名：ShowAllType
'作  用：显示所有文章类型和简介（栏目导航）
'参  数：无
'=================================================
dim  sqlnbTyp,rsnbTyp
sqlnbTyp="select * from StuZoneType order by id"
Set rsnbTyp= Server.CreateObject("ADODB.Recordset")
rsnbTyp.open sqlnbTyp,conn,1,1
	if rsnbTyp.bof and rsnbTyp.eof then 
		response.Write "&nbsp;没有任何类别"
	else
		dim sqlC,rsC,strCName
		rsnbTyp.movefirst
		do while not rsnbTyp.eof
			strCName= "<div class='rbigclass'><a href='stuZone.asp?nbtyped=" & Server.URLEncode(rsnbTyp("nbtyped")) & "'>" & rsnbTyp("nbtyped") & "</a></div><div class='rBrief'>" & replace(rsnbTyp("Brief"),"../","./")  & "</div>"
		response.write strCName & "<p></p>"
		rsnbTyp.movenext
		loop
		rsC.close
		set rsC=nothing
	  rsnbTyp.close
      set rsnbTyp=nothing
	end if
 %>
          
          
          </div>
        </div>
        
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
