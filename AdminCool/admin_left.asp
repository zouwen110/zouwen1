<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>管理页面</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {
	 BACKGROUND: #EEF2FB; MARGIN: 0px; FONT: 12px 宋体; 
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px 宋体
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px 宋体; COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #428eff; TEXT-DECORATION: underline
}
.sec_menu {
	 BACKGROUND: #EEF2FB; OVERFLOW: hidden;
}
.menu_title {
FONT-WEIGHT: bold;
}
.menu_title SPAN {
	FONT-WEIGHT: bold; POSITION: relative; TOP: 2px
}
.menu_title2 {
FONT-WEIGHT: bold;
}
.menu_title a:link{
	font-weight:bold;}
.menu_title a:visited{
	font-weight:bold;}
.menu_title2 a:link{
	font-weight:bold;}
.menu_title2 a:visited{
	font-weight:bold;}	
.menu_title2 SPAN {
	FONT-WEIGHT: bold;COLOR: #006600; POSITION: relative; TOP: 2px
}
.MM {
	width: 182px;
	height:26px;
	margin: 0px;
	padding: 0px;
	left: 0px;
	top: 0px;
	right: 0px;
	bottom: 0px;
	clip: rect(0px,0px,0px,0px);
	background-image: url(images/menu_bg1.gif);
}
.MM a:link {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	color: #333333;
	background-image: url(images/menu_bg1.gif);
	background-repeat: no-repeat;
	height: 26px;
	width: 182px;
	display: block;
	text-align: center;
	margin: 0px;
	padding: 0px;
	overflow: hidden;
	text-decoration: none;
}
.MM a:active {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	color: #333333;
	background-image: url(images/menu_bg.gif);
	background-repeat: no-repeat;
	height: 26px;
	width: 182px;
	display: block;
	text-align: center;
	margin: 0px;
	padding: 0px;
	overflow: hidden;
	text-decoration: none;
}
.MM a:hover {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	font-weight: bold;
	color: #006600;
	background-image: url(images/menu_bg2.gif);
	background-repeat: no-repeat;
	text-align: center;
	display: block;
	margin: 0px;
	padding: 0px;
	height: 26px;
	width: 182px;
	text-decoration: none;
}
</STYLE>

<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<script language=JavaScript>
function logout(){
	if (confirm("您确定要退出后台管理系统吗？"))
	top.location = "logout.asp";
	return false;
}
</script>
<META content="MSHTML 6.00.3790.2817" name=GENERATOR>
<BODY leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=left border=0>
  <TBODY>
  <TR>
    <TD vAlign=top>&nbsp;
<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
          <TR>
            <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(220)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN><a href="right.asp" target="main" >后台首页</a></SPAN> </div></TD>
          </TR>
        </TBODY>
	    </TABLE>
      <%If logr() Then %>
	  
	  <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(0)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>基本设置</SPAN></div></TD>
        </TR>
        <TR>
          <TD id=submenu0 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 182px">
            <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
              <TBODY>
              <TR>
                <TD height=5 background="images/menu_topline.gif"></TD>
              </TR>
              <TR>
                <TD  class="MM" ><div align="center"><A
                  href="web_settings.asp"
                  target=main>网站信息设置</A></div></TD>
              </TR>
              <TR>
                <TD  class="MM" ><div align="center"><A
                  href="web_Advanced.asp"
                  target=main>网站高级设置</A></div></TD>
              </TR>             <TR>
                <TD class="MM"  ><div align="center"><A
                  href="admin_list.asp"
                  target=main>后台用户管理</A></div></TD>
             </TR>              

              <TBODY></TBODY></TABLE>
            </DIV>
</TD></TR></TBODY></TABLE>
     
	  
</div>

<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(91)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>主题管理</SPAN> </div></TD>
        </TR>
        <TR>
          <TD id=submenu91 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 182px">
            <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
              <TBODY>
              <TR>
                <TD height=5 background="images/menu_topline.gif"></TD></TR>
				<TR>
                 <TD  class="MM" > <div align="center"><A
                  href="Theme_add.asp"
                  target=main>添加主题</A></div></TD></TR>
				   <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="ThemeSetting.asp"
                  target=main>主题列表</A></div></TD></TR>
              <TBODY></TBODY></TABLE>
            </DIV> </TD></TR></TBODY></TABLE>
<table cellspacing="0" cellpadding="0" width="182" align="center">
  <tbody>
    <tr>
      <td class="menu_title" id="menuTitle1"
          onmouseover="this.className='menu_title2';" onClick="showsubmenu(71)"
          onmouseout="this.className='menu_title';"
          background="images/menu_bgs.gif"
            height="30"><div align="center"><span>模板管理</span> </div></td>
    </tr>
    <tr>
      <td id="submenu71" style="DISPLAY: none"><div class="sec_menu" style="WIDTH: 182px">
        <table cellspacing="0" cellpadding="0" width="182" align="center">
          <tbody>
            <tr>
              <td height="5" background="images/menu_topline.gif"></td>
            </tr>
          </tbody>
          <tr>
            <td  class="MM" ><div align="center"><a
                  href="models_type_list.asp"
                  target="main">模板分类管理</a></div></td>
          </tr>
          <tr>
            <td  class="MM" ><div align="center"><a
                  href="web_models_add.asp"
                  target="main">添加模板</a></div></td>
          </tr>
          <tr>
            <td  class="MM" ><div align="center"><a
                  href="web_models.asp"
                  target="main">模板列表</a></div></td>
          </tr>
          <tbody>
          </tbody>
        </table>
      </div></td>
    </tr>
  </tbody>
</table>
<%End If %>
<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(1)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>导航管理</SPAN> </div></TD>
        </TR>
        <TR>
          <TD id=submenu1 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 182px">
            <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
              <TBODY>
              <TR>
                <TD height=5 background="images/menu_topline.gif"></TD></TR>
				<TR>
                 <TD  class="MM" > <div align="center"><A
                  href="menu_type_list.asp"
                  target=main>导航分类</A></div></TD></TR>
				   <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="menu_add.asp"
                  target=main>添加导航</A></div></TD></TR><TR>
                 <TD  class="MM" > <div align="center"><A
                  href="menu_list.asp"
                  target=main>导航列表</A></div></TD>
		   </TR>
              <TBODY></TBODY></TABLE>
            </DIV> </TD></TR></TBODY></TABLE>
				
<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
   <TBODY>
     <TR>
       <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(2)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>栏目管理</SPAN> </div></TD>
     </TR>
     <TR>
       <TD id=submenu2 style="DISPLAY: none"><DIV class=sec_menu style="WIDTH: 182px">
           <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
             <TBODY>
               <TR>
                 <TD height=5 background="images/menu_topline.gif"></TD>
               </TR>
               <TR>
                  <TD  class="MM" > <div align="center"><A
                  href="category_add.asp?ppid=1"
                  target=main>添加一级栏目</A></div></TD>
               </TR>
               <TR>
                  <TD  class="MM" > <div align="center"><A
                  href="category_list.asp"
                  target=main>栏目列表</A></div></TD>
               </TR>
             <TBODY>
             </TBODY>
           </TABLE>
       </DIV> </TD>
     </TR>
   </TBODY>
 </TABLE>
 <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(4)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>文章管理</SPAN> </div></TD>
        </TR>
        <TR>
          <TD id=submenu4 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 182px">
            <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
              <TBODY>
              <TR>
                <TD height=5 background="images/menu_topline.gif"></TD></TR>
		<TR>
                 <TD  class="MM" > <div align="center"><A
                  href="article_add.asp"
                  target=main>添加文章</A></div></TD>
		</TR>
             <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="article_list.asp"
                  target=main>文章列表</A></div></TD>
             </TR>
             <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="comment_list.asp"
                  target=main>文章评论</A></div></TD>
             </TR>
             <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="keywords_list.asp"
                  target=main>文章关键词</A></div></TD>
             </TR>
             <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="author_list.asp"
                  target=main>文章来源</A></div></TD>
             </TR>				 	
              <TBODY></TBODY></TABLE>
            </DIV> </TD></TR></TBODY></TABLE>

<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
   <TBODY>
     <TR>
       <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(222)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>文件管理</SPAN> </div></TD>
     </TR>
     <TR>
       <TD id=submenu222 style="DISPLAY: none"><DIV class=sec_menu style="WIDTH: 182px">
           <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
             <TBODY>
               <TR>
                 <TD height=5 background="images/menu_topline.gif"></TD>
               </TR>
               <TR>
                  <TD  class="MM" > <div align="center"><A
                  href="File_Setup.asp"
                  target=main>上传设置</A></div></TD>
               </TR>
               <TR>
                  <TD  class="MM" > <div align="center"><A
                  href="File_Upload.asp"
                  target=main>上传文件</A></div></TD>
               </TR>			   
               <TR>
                  <TD  class="MM" > <div align="center"><A
                  href="File_List.asp"
                  target=main>文件列表</A></div></TD>
               </TR>
             <TBODY>
             </TBODY>
           </TABLE>
       </DIV> </TD>
     </TR>
   </TBODY>
 </TABLE>
 
 		<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
          <TBODY>
            <TR>
              <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(21)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>相册管理</SPAN> </div></TD>
            </TR>
            <TR>
              <TD id=submenu21 style="DISPLAY: none"><DIV class=sec_menu style="WIDTH: 182px">
                  <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
                    <TBODY>
                      <TR>
                        <TD height=5 background="images/menu_topline.gif"></TD>
                      </TR>
                      <TR>
                         <TD  class="MM" > <div align="center"><A href="ad_position_add.asp"  target=main>添加新相册</A></div></TD>
                      </TR>
                      <TR>
                         <TD  class="MM" > <div align="center"><A href="ad_position_list.asp"  target=main>相册列表</A></div></TD>
                      </TR>
                      <TR>
                         <TD  class="MM" > <div align="center"><A href="ad_MutiAdd.asp"  target=main>添加图片</A></div></TD>
                      </TR>						  				  					  
                      <TR>
                         <TD  class="MM" > <div align="center"><A href="ad_list.asp"  target=main>图片列表</A></div></TD>
                      </TR>
                    <TBODY>
                    </TBODY>
                  </TABLE>
              </DIV> </TD>
            </TR>
          </TBODY>
        </TABLE>
 		<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
                      <TBODY>
                        <TR>
                          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(18)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>留言管理</SPAN> </div></TD>
                        </TR>
                        <TR>
                          <TD id=submenu18 style="DISPLAY: none"><DIV class=sec_menu style="WIDTH: 182px">
                              <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
                                <TBODY>
                                  <TR>
                                    <TD height=5 background="images/menu_topline.gif"></TD>
                                  </TR>
                                  <TR>
                                     <TD  class="MM" > <div align="center"><A  href="message_list.asp" target=main>留言列表</A></div></TD>
                                  </TR>
                                <TBODY>
                                </TBODY>
                              </TABLE>
                          </DIV> </TD>
                        </TR>
                      </TBODY>
        </TABLE>
					<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(16)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>友情链接</SPAN> </div></TD>
        </TR>
        <TR>
          <TD id=submenu16 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 182px">
            <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
              <TBODY>
              <TR>
                <TD height=5 background="images/menu_topline.gif"></TD></TR>
              <TR>
                 <TD  class="MM" > <div align="center"><A href="link_add.asp"  target=main>添加友情链接</A></div></TD></TR>
 <TR>
                 <TD  class="MM" > <div align="center"><A href="link_list.asp"  target=main>友情链接列表</A></div></TD></TR>
              <TBODY></TBODY></TABLE>
            </DIV> </TD></TR></TBODY></TABLE>
	
				

					<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(126)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>数据库管理</SPAN> </div></TD>
        </TR>
        <TR>
          <TD id=submenu126 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 182px">
            <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
              <TBODY>
              <TR>
                <TD height=5 background="images/menu_topline.gif"></TD></TR>
				<TR>
                 <TD  class="MM" > <div align="center"><A
                  href="Data_Backup.asp"
                  target=main>备份数据库</A></div></TD></TR>
				   <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="Data_Restore.asp"
                  target=main>还原数据库</A></div></TD></TR>
				   <TR>
                 <TD  class="MM" > <div align="center"><A
                  href="Data_List.asp"
                  target=main>备份数据列表</A></div></TD></TR>					 
              <TBODY></TBODY></TABLE>
            </DIV> </TD></TR></TBODY></TABLE>	
			
				    <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
                  <TBODY>
                    <TR>
                      <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(33)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>生成静态</SPAN> </div></TD>
                    </TR>
                    <TR>
                      <TD id=submenu33 style="DISPLAY: none"><DIV class=sec_menu style="WIDTH: 182px">
                          <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
                            <TBODY>
                              <TR>
                                <TD height=5 background="images/menu_topline.gif"></TD>
                              </TR>
                              <TR>
                                 <TD  class="MM" > <div align="center"><A href="html_index.asp"  target=main>生成首页</A></div></TD>
                              </TR>
							  <TR>
                                 <TD  class="MM" > <div align="center"><a href="html_items.asp" target="main">生成栏目</a></div></TD>
                              </TR>
							  <TR>
                                 <TD  class="MM" > <div align="center"><A href="html_article.asp"  target=main>生成内容</A></div></TD>
                              </TR>
                            <TBODY>
                            </TBODY>
                          </TABLE>
                      </DIV> </TD>
                    </TR>
                  </TBODY>
	    </TABLE>
								
				<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(20)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN><a href="#" target="_self" onClick="logout();">退出登录</a></SPAN> </div></TD>
        </TR></TBODY></TABLE>


<TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
                  <TBODY>
                    <TR>
                      <TD class=menu_title id=menuTitle1
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(333)
          onmouseout="this.className='menu_title';"
          background=images/menu_bgs.gif
            height=30><div align="center"><SPAN>版权信息</SPAN> </div></TD>
                    </TR>
                    <TR>
                      <TD id=submenu333 ><DIV class=sec_menu style="WIDTH: 182px">
                          <TABLE cellSpacing=0 cellPadding=0 width=182 align=center>
                            <TBODY>
                              <TR>
                                <TD height=5 background="images/menu_topline.gif"></TD>
                              </TR>
                              <TR>
                                <TD height="40"  ><div align="center">
                                  <p><A href="http://www.58yingxiao.net/"  target="_blank" title="58网络营销培训网">58网络营销培训网</A> 版权所有</p>
                                </div></TD>
                              </TR>

                            <TBODY>
                            </TBODY>
                          </TABLE>
                      </DIV> </TD>
                    </TR>
                  </TBODY>
	    </TABLE>		 

    </TR></TBODY></TABLE></BODY></HTML>
<%Call DbconnEnd()
%>