<!-- #include file="../../../include/config.asp" -->
<!-- #include file="../../../include/function.asp" -->
<!-- #include file="../../../languages/en-us.asp" -->
<%
url_page = "article"
site_title = "新秀免费企业网站系统 sinsiu 1.2 beta1"
channel_title = "资讯中心"
cat_id = 0
id = 8
url = "article_8.html"
suffix = "html"
page_title = "ART1111"
keywords = "网站关键字"
describe = "网站描述"
cat_id = "2"
cat_name = "文章二级分类"
dim nav(7,2)
nav(0,0) = "首页"
nav(0,1) = "/test/likai/sinsiu_v1.2_beta1/?index.html"
nav(0,2) = "0"
nav(1,0) = "关于我们"
nav(1,1) = "/test/likai/sinsiu_v1.2_beta1/?about.html"
nav(1,2) = "0"
nav(2,0) = "产品展示"
nav(2,1) = "/test/likai/sinsiu_v1.2_beta1/?product.html"
nav(2,2) = "0"
nav(3,0) = "资讯中心"
nav(3,1) = "/test/likai/sinsiu_v1.2_beta1/?article.html"
nav(3,2) = "0"
nav(4,0) = "人才招聘"
nav(4,1) = "/test/likai/sinsiu_v1.2_beta1/?recruit.html"
nav(4,2) = "0"
nav(5,0) = "下载中心"
nav(5,1) = "/test/likai/sinsiu_v1.2_beta1/?download.html"
nav(5,2) = "0"
nav(6,0) = "留言板"
nav(6,1) = "/test/likai/sinsiu_v1.2_beta1/?message.html"
nav(6,2) = "0"
nav(7,0) = "新秀官网"
nav(7,1) = "http://www.sinsiu.com"
nav(7,2) = "1"
nav_u = 7
dim nav_name(2)
nav_name(0) = "word"
nav_name(1) = "url"
nav_name(2) = "target"
dim focus(1,3)
focus(0,0) = "/test/likai/sinsiu_v1.2_beta1/images/focus_1.jpg"
focus(0,1) = "index.asp"
focus(0,2) = "焦点图片"
focus(0,3) = "default"
focus(1,0) = "/test/likai/sinsiu_v1.2_beta1/images/focus_2.jpg"
focus(1,1) = "index.asp"
focus(1,2) = "焦点图片"
focus(1,3) = "default"
focus_u = 1
dim focus_name(2)
focus_name(0) = "path"
focus_name(1) = "url"
focus_name(2) = "title"
dim tree(1,3)
tree(0,0) = "1"
tree(0,1) = "0"
tree(0,2) = "文章一级分类"
tree(0,3) = "1"
tree(1,0) = "2"
tree(1,1) = "1"
tree(1,2) = "文章二级分类"
tree(1,3) = "2"
tree_u = 1
dim tree_name(3)
tree_name(0) = "id"
tree_name(1) = "parent"
tree_name(2) = "name"
tree_name(3) = "grade"
show_sheet = 0
dim article(0,4)
article(0,0) = "8"
article(0,1) = "ART1111"
article(0,2) = "DFD"
article(0,3) = "2013/8/23 21:03:00"
article(0,4) = "SDGDSGDSART1111111111111111111111111111"
article_u = 0
dim article_name(4)
article_name(0) = "art_id"
article_name(1) = "art_title"
article_name(2) = "att_author"
article_name(3) = "art_add_time"
article_name(4) = "art_text"
dim footer_nav(4,2)
footer_nav(0,0) = "首页"
footer_nav(0,1) = "/test/likai/sinsiu_v1.2_beta1/?index.html"
footer_nav(0,2) = "0"
footer_nav(1,0) = "关于我们"
footer_nav(1,1) = "/test/likai/sinsiu_v1.2_beta1/?about.html"
footer_nav(1,2) = "0"
footer_nav(2,0) = "产品展示"
footer_nav(2,1) = "/test/likai/sinsiu_v1.2_beta1/?product.html"
footer_nav(2,2) = "0"
footer_nav(3,0) = "资讯中心"
footer_nav(3,1) = "/test/likai/sinsiu_v1.2_beta1/?article.html"
footer_nav(3,2) = "0"
footer_nav(4,0) = "人才招聘"
footer_nav(4,1) = "/test/likai/sinsiu_v1.2_beta1/?recruit.html"
footer_nav(4,2) = "0"
footer_nav_u = 4
dim footer_nav_name(2)
footer_nav_name(0) = "word"
footer_nav_name(1) = "url"
footer_nav_name(2) = "target"
sit_code = "粤ICP备12345678号"
sit_code_url = "http://www.sinsiu.com"
sit_tec = "新秀工作室"
sit_tec_url = "http://www.sinsiu.com"
sit_count_code = "统计代码"
service = "在线客服"
%>
<!-- #include file="../compile/article.asp" -->
