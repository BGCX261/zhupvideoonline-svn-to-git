<%response.end%>
<div id="header">
  <div class="logo"><a href="{$S_ROOT}"><img src="{$S_ROOT}images/logo.jpg" /></a></div>
  <div class="search">
    <form method="post" action="{$S_ROOT}index.asp?search.{$suffix}">
      <input class="text" name="search" type="text" maxlength="30" />
      {if:$S_LANG,neq,"en-us"}
      <input class="button" name="submit" type="submit" value="{$L_SEARCH}"/>
      {else}
      <input class="button" name="submit" type="submit" style="width:35px;font-size:10px;font-weight:normal;" value="{$L_SEARCH}"/>
      {/if}
    </form>
  </div>
  <div id="nav">
    <ul>
      {for:$i,0,$nav_u}
      <li {if:$i,eq,0}class="first"{/if}>
	  	<a href="{$nav[$i]['url']}" {if:$nav[$i]['target'],eq,1}target="_blank"{/if}>{$nav[$i]['word']}</a>
		{*按中文名称判断下拉菜单，不用建议删除*}
		{$word=$nav[$i]['word']}
		{if:$word,eq,"关于我们"}
			{*标签移到这里*}
		{elseif:$word,eq,"产品展示"}
			{*标签移到这里*}
		{elseif:$word,eq,"资讯中心"}
			{*标签移到这里*}
		{/if}
		
		{*按顺序判断下拉菜单，序号$i从0开始*}
	  	{if:$i,eq,1}
			<ul>
				{for:$j,0,$nav_about_u}
				<li><a href="{$nav_about[$j]['url']}">{$nav_about[$j]['word']}</a></li>
				{/for}
			</ul>
		{elseif:$i,eq,2}
			<ul>
			  {for:$j,0,$nav_pro_u}
				{if:$nav_pro[$j]['grade'],eq,1}
				<li><a href="{$S_ROOT}?product_{$nav_pro[$j]['id']}_0_0.{$suffix}">{$nav_pro[$j]['name']}</a></li>
				{/if}
			  {/for}
			</ul>		
		{elseif:$i,eq,3}
			<ul>
			  {for:$j,0,$nav_art_u}
				{if:$nav_art[$j]['grade'],eq,1}
				<li><a href="{$S_ROOT}?article_{$nav_art[$j]['id']}_0_0.{$suffix}">{$nav_art[$j]['name']}</a></li>
				{/if}
			  {/for}
			</ul>		
		{/if}
	  </li>
	  {/for}
      <div class="clear"></div>
    </ul>
  </div>
</div>
<!-- 新秀 -->
