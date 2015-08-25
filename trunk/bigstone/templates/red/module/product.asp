<%response.end%>
<div class="box2 list" id="product">
  <div class="head">
    <div class="l"></div>
    <span>{$L_NEW_PRODUCT}</span>
    <a class="more" href="{$S_ROOT}?product.{$suffix}">{$L_MORE}</a>
    <div class="r"></div>
  </div>
  <ul class="main">
    {for:$i,0,$product_u}
	<li><a href="{$S_ROOT}?product_{$product[$i]['art_id']}.{$suffix}" target="_blank">{$cut_str($product[$i]['art_title'],14)}</a></li>
    {/for}
  </ul>
</div>
<!-- 新秀 -->
