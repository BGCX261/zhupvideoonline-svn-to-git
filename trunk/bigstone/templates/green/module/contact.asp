<%response.end%>
<div class="box" id="contact">
  <div class="head"><div class="l"></div><span>{$L_CONTACT}</span><div class="r"></div></div>
  <div class="main">
    {for:$i,0,$contact_u}
    <span>{$contact[$i]['cue']}：</span>{$contact[$i]['content']}<br />
    {/for}
  </div>
</div>
<!-- 新秀 -->
