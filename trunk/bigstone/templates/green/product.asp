<%response.end%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta name ="keywords" content="{$keywords}" />
<meta name ="description" content="{$describe}" />
<title>{if:$page_title,neq,""}{$page_title} - {/if}{if:$cat_name,neq,""}{$cat_name} - {/if}{$channel_title} - {$site_title}</title>
<link href="{$S_TPL_PATH}css.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="{$S_TPL_PATH}js/script.js"></script>
<!--[if IE 6]>
<script type="text/javascript" src="{$S_TPL_PATH}js/DD_belatedPNG_0.0.8a.js"></script>
<script type="text/javascript">
DD_belatedPNG.fix("*");
</script>
<![endif]-->
</head>
<body>
  {inc:module/header}
  {inc:module/focus}
  <div id="main">
    <div id="left">
      {inc:module/product_tree}
    </div>
    <div id="right2">
      {inc:module/here}
      {inc:module/product_main}
    </div>
    <div class="clear"></div>
  </div>
  {inc:module/footer}
  {inc:module/service}
</body>
</html>