<%response.end%>
<div class="box2" id="about">
  <div class="head">
    <div class="l"></div>
    <span>{if:$about_title,egt,""}{$about_title}{else}{$L_ABOUT_US}{/if}</span>
    <a class="more" href="{$S_ROOT}?about.{$suffix}">{$L_MORE}</a>
    <div class="r"></div>
  </div>
  <div class="main">
    <div class="img"><img src="{$S_TPL_PATH}images/about.jpg" /></div>
    {$about_text}
    <a class="more" href="{$S_ROOT}?about.{$suffix}">【{$L_DETAILED}】</a>
  </div>
</div>
<!-- 新秀 -->
