<%response.end%>
<div class="box" id="contact">
  <div class="head"><span>{$L_CONTACT}</span></div>
  <div class="main">
    {for:$i,0,$contact_u}
    <span>{$contact[$i]['cue']}：</span>{$contact[$i]['content']}<br />
    {/for}
  </div>
</div>
<!-- 新秀 -->
