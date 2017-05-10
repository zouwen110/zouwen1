<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=7">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../inc/web_config.asp" -->
<!-- #include file="../inc/html_clear.asp" -->
<%
search_q=request.querystring("q")
%>
<title>搜寻：<%=search_q%> - 58网络营销培训博客-欢迎您的光临</title>
<link href="/css/LowerPriority/style.css" rel="stylesheet" type="text/css" media="screen" />
<script  language="javascript" src="/js/slidealbum.js"></script>
</head>
<body>
<%
keywords=split(search_q," ")
c=ubound(keywords)
for i=0 to c
if i=0 then
search_sql1=search_sql1&"where  ( [title] like '%"&keywords(i)&"%'"
keywords_all=keywords(i)
else
search_sql1=search_sql1&" or   [title] like '%"&keywords(i)&"%'"
keywords_all=keywords_all&"+"&keywords(i)
end if
next

s_sql="select [title],[content],[file_path],[time] from [article] "&search_sql1&" )  and view_yes=1 order by [time] desc"
%>
<div id="wrapper">
	<div id="header">
    <div class="logoBG">
		<div id="logo">
        <div class="left">
		<h1><a href="/" title="58网络营销培训博客-欢迎您的光临">58网络营销培训博客-欢迎您的光临</a></h1>		
		<p id="slogan">打造全国优秀免费网络营销培训基地，记住我们的网址www.58yingxiao.net</p>	
                    </div>
		<div id="search">
			<form method="get" action="/query/index.asp">
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='请输入关键词';" 
onfocus="if(this.value=='请输入关键词') this.value='';" value="请输入关键词" /><input type="submit" id="search-submit" value="搜 寻" />
			</form>
		</div></div>
        
		</div><div class="clearfix"></div>
        <div class="menuBG"><div id="menu">
<ul><li><a href='/' >首 页</a></li> <li><a href='/blog/' >博 客</a></li> <li><a href='/album/' >相 册</a></li> <li><a href='/Category/About' >关 于</a></li> <li><a href='/FeedBack/' >留 言</a></li> <li><a href='/Category/Contact' >联 系</a></li> <li><a href='http://www.58daohang.net' target='_blank'>官方网站</a></li> </ul>	</div>
                </div>
	</div>
	<div id="page">
		<div id="content">
			<div class="Web_Position">您现在的位置: <a href="/">首页</a> > <a href='/query/'>搜寻</a></div>
		<div class="clearfix"></div>		
<!--search content start-->
<div id="search_content" class="clearfix">

<%
if search_q<>"" then 

set rs=server.createobject("adodb.recordset")
rs.open(s_sql),cn,1,1
%>

<%'=============分页定义开始，要放在数据库打开之后
if err.number<>0 then '错误处理
response.write "数据库操作失败：" & err.description
err.clear
else
if not (rs.eof and rs.bof) then '检测记录集是否为空
r=cint(rs.RecordCount) '记录总数
rowcount = 10 '设置每一页的数据记录数，可根据实际自定义
rs.pagesize = rowcount '分页记录集每页显示记录数
maxpagecount=rs.pagecount '分页页数
page=request.querystring("page")
  if page="" then
  page=1
  end if
rs.absolutepage = page 
rcount1=0
pagestart=page-5
pageend=page+5
if pagestart<1 then
pagestart=1
end if
if pageend>maxpagecount then
pageend=maxpagecount
end if
rcount=rs.RecordCount
'=============分页定义结束%>

<!--position start-->
<div class="searchtip">您正在搜寻"<span class="FontRed"><%=search_q%></span>",找到相关信息 <span class="font_brown"><%=rcount%></span> 条</div>
<!--position end-->
<!--list start-->
<div class="result_list">
<div class="gray">提示：用空格隔开多个搜寻关键词可获取更理想结果，如'李毅 足球'。</div>
<dl>

<%'===========循环体开始
do while not rs.eof and rowcount%>
<%
title1=left(rs("title"),30)
for i=0 to c
title1=Replace(title1, keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next

content1=left(Clearhtml(rs("content")),110)
for i=0 to c
content1=Replace(content1,keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next
%>
<dt ><a href='<%="/"&Article_FolderName&"/"&rs("file_path")%>' target='_blank' title='<%=rs("title")%>'><%=title1%></a></dt>
<dd><%=content1%>...</dd>
<dd class="font12 arial font_green line"><a href='<%="/"&Article_FolderName&"/"&rs("file_path")%>' target='_blank'><span class="font_green"><%=web_url&Article_FolderName&"/"&rs("file_path")%></span></a><%=year(rs("time"))%>-<%=month(rs("time"))%>-<%=day(rs("time"))%></dd>
<%
rowcount=rowcount-1 
rs.movenext
loop
 '===========循环体结束%>

</dl>
</div>
<!--list end-->

<!--page start-->
<div class="result_page clearfix">
<!--#include file="../inc/page_list.asp"-->
</div>
<!--page end-->

<%
else
response.write "<div class='search_welcome'>很抱歉,没有找到与 <span class='FontRed'>"&search_q&"</span> 相关的信息！<p >提示：用空格隔开多个搜寻关键词可获取更理想结果，如'足球 李毅'。</p></div>"
end if
end if
end if%>
</div>
<!--search content end-->	
		<div style="clear: both;">&nbsp;</div>
		</div>
		<div id="sidebar">
			<ul>	
				<li>
					<h2>个人档案</h2>
					<p class="myphoto"><img src="/images/up_images/20121223234645.jpg" width="80" height="89"></p>
					<p class="myintro">杨磊<br>2012年06月08日<br>中国河北石家庄<br>279018860@qq.com</p>
					<p class="clearfix">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;希望这个平台能和更多志同道合的朋友一起探讨和交流，让我们一起走向辉煌</a></p>
					</li>
				<li>
					<h2>精彩相册</h2>
<!--slide album start-->
<div id="slidebox">
	<div id="slider">
<div class='slide'><A href='/Gallery/64/' target='_blank' title='杨磊'><img class='diapo' src='/photos/20121223234246959.jpg' alt='杨磊' width='210' ></a><div class='text'><span class='title'>杨磊</span></div></div><div class='slide'><A href='/Gallery/64/' target='_blank' title='站长杨磊近照'><img class='diapo' src='/photos/20121223233953817.jpg' alt='站长杨磊近照' width='210' ></a><div class='text'><span class='title'>站长杨磊近照</span></div></div><div class='slide'><A href='/Gallery/63/' target='_blank' title='世界风光图片'><img class='diapo' src='/photos/20120529123656316.jpg' alt='世界风光图片' width='210' ></a><div class='text'><span class='title'>世界风光图片</span></div></div><div class='slide'><A href='/Gallery/63/' target='_blank' title='世界风光图片'><img class='diapo' src='/photos/20120529123656141.jpg' alt='世界风光图片' width='210' ></a><div class='text'><span class='title'>世界风光图片</span></div></div>
	</div>
<script type="text/javascript">
/* ==== start script ==== */
slider.init();
</script>
</div>
<!--slide album end-->
				</li>
				<li>
					<h2>博客分类</h2>
<ul><li><a href='/Category/Enterntainment/'>娱乐资讯</a> (0) <a href='/rss/Feed.asp?CatID=133' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Love/'>情感天地</a> (0) <a href='/rss/Feed.asp?CatID=135' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Internet/'>互联网</a> (0) <a href='/rss/Feed.asp?CatID=136' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Favorite/'>个人收藏</a> (0) <a href='/rss/Feed.asp?CatID=138' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Diary/'>个人日记</a> (0) <a href='/rss/Feed.asp?CatID=137' target='_blank'><img src='/images/rss_icon.gif'></a></li></ul>
				</li>
				<li>
					<h2>热门文章</h2>
暂无信息。
				</li>
			</ul>
		</div>
		<br class="clearfix" />
	</div>
</div>
<div id="footer">
	<p>Copyright &copy; 2012 58网络营销培训博客(<A href="http://www.58daohang.net/" target=_blank>Hitux.com</A>) Britar Yao All rights reserved<BR>&nbsp;Powered by <A href="www.58daohang.net" target=_blank>HituxBlog V1.4</A> <img src="/images/hituxblog-logo.png" width="80" alt='HituxBlog V1.4'> <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss/feed.xml" target="_blank"><img src="/images/xml_icon.gif"></a> Themes By <A href="http://www.58daohang.net/" target=_blank>FreeCssTemplates</a>
<p>58网络营销培训博客致力于<a href="http://www.hitux.com/">seo网站优化免费培训</a>、<a href="http://www.hitux.com/">网赚项目免费培训</a>、<a href="http://www.hitux.com/">网上创业免费培训</a>等业务。联系QQ：2528955292</p></p>
</div>
</body>
</html>

