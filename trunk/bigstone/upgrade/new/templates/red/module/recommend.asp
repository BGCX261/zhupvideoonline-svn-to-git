<%response.end%>
<div class="box2" id="recommend">
  <div class="head">
    <div class="l"></div>
    <span>{$L_RECOMMEND}</span>
    <a class="more" href="{$S_ROOT}?product.{$suffix}">{$L_MORE}</a>
    <div class="r"></div>
  </div>
  <div class="main">
  <!-------------------------->
  <div id="roll">
    <div id="roll_sheet" onmouseover="over_roll()" onmouseout="out_roll()">
      {for:$i,0,recommend_u}
      <div name="roll_unit">
        <table cellspacing="0" cellpadding="0">
          <tr><td class="img">
            <a href="{$S_ROOT}?product_{$recommend[$i]['art_id']}.{$suffix}" target="_blank"><img src="{$S_ROOT}{$recommend[$i]['art_reduce_img']}" onload="picresize(this,150,150)"/></a>
          </td></tr>
          <tr><td class="title">
            <a href="{$S_ROOT}?product_{$recommend[$i]['art_id']}.{$suffix}" target="_blank">{$cut_str($recommend[$i]['art_title'],11)}</a>
          </td></tr>
        </table>
      </div>
      {/for}
    </div>
  </div>
  <!-------------------------->
  </div>
</div>
<script type="text/javascript" src="{$S_TPL_PATH}js/roll.js"></script>
<!-- 新秀 -->
