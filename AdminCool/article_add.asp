<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/rand.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/article_to_html.asp" -->
<!-- #include file="../inc/x_to_html/blog_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/blog_class_list_to_html.asp" -->

<% '�������ݵ����ݱ�
act=Request("act")
If act="save" Then 

a_title=request.form("a_title")
a_cid=trim(request.form("cid"))
a_pid=trim(request.form("pid"))
a_ppid=trim(request.form("ppid"))
a_url=trim(request.form("a_url"))
a_image=trim(request.form("web_image"))
a_keywords=trim(request.form("a_keywords"))
a_description=trim(request.form("a_description"))
a_content=request.form("a_content")
a_from_name=trim(request.form("a_from_name"))
a_from_url=trim(request.form("a_from_url"))
a_author=trim(request.form("a_author"))
a_hit=trim(request.form("a_hit"))
a_index_push=trim(request.form("a_index_push"))
a_keywords_yes=trim(request.form("a_keywords_yes"))
a_time=now()


'�滻�ؼ���

'�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=11"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Search_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

if a_keywords_yes=1 then
set rs_1=server.createobject("adodb.recordset")
sql="select [web_theme] from web_settings "
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Theme_Folder=rs_1("web_theme")
end if
rs_1.close
set rs_1=nothing

set rs=server.createobject("adodb.recordset")
sql="select [name],[url] from web_keywords order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof  then
do while not rs.eof 

if rs("url")="" then
name_url="/"&Search_FolderName&"/?q="&rs("name")
else
name_url=rs("url")
end if

a_content=replace(a_content,rs("name"),"<a href='"&name_url&"' target='_blank'>"&rs("name")&"</a>")
rs.movenext
loop
end if
rs.close
set rs=nothing
end if

a_content=replace(a_content,"<IMG ","<IMG alt='"&a_title&"' ")

set rs=server.createobject("adodb.recordset")
sql="select * from article"
rs.open(sql),cn,1,3
rs.addnew
rs("title")=a_title
rs("cid")=a_cid
rs("pid")=a_pid
rs("ppid")=a_ppid
rs("url")=a_url
rs("image")=a_image
rs("keywords")=a_keywords
rs("description")=a_description
rs("content")=a_content
rs("from_name")=a_from_name
rs("from_url")=a_from_url
rs("author")=a_author
rs("hit")=a_hit
rs("index_push")=a_index_push
rs("time")=a_time
rs("edit_time")=a_time
if request.form("slide_yes")<>"" then
rs("slide_yes")=1
end if
rs("File_Path")=a7&minute(now)&second(now)&".html"
rs.update
rs.close
set rs=nothing
%>

<% '������ҳ
call blog_index_to_html()
call index_to_html()
%>

<% '�������¾�̬ҳ
set rs2=server.createobject("adodb.recordset")
sql="select top 1 [id],[title],file_path from [article] where [title]='"&a_title&"' order by [time] desc"
rs2.open(sql),cn,1,1
if not rs2.eof  then
a_id=rs2("id")
a_title=rs2("title")
a_link=Article_FolderName&"/"&rs2("file_path")
call article_to_html(a_id)
'�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=9"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
if rs_1("FolderName")<>"" then
Article_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing
end if
%>



<%
response.Write "<script language='javascript'>alert('���ӳɹ���');location.href='article_list.asp';</script>"
end if 

 %>
<script type="text/javascript" charset="utf-8" src="../KKKeditor/kindeditor.js"></script>
<script type="text/javascript" src="../KKKeditor/editor.js"></script>
 <!-- ���������˵� ��ʼ -->
<script language="JavaScript">
<!--
<%
'�������ݱ��浽����
Dim count2,rsClass2,sqlClass2
set rsClass2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from category where ppid=2  and ClassType=1 order by id " 
rsClass2.open sqlClass2,cn,1,1
%>
var subval2 = new Array();
//����ṹ��һ����ֵ,������ֵ,������ʾֵ
<%
count2 = 0
do while not rsClass2.eof
%>
subval2[<%=count2%>] = new Array('<%=rsClass2("pID")%>','<%=rsClass2("ID")%>','<%=rsClass2("Name")%>')
<%
count2 = count2 + 1
rsClass2.movenext
loop
rsClass2.close
%>

<%
'�������ݱ��浽����
Dim count3,rsClass3,sqlClass3
set rsClass3=server.createobject("adodb.recordset")
sqlClass3="select id,pid,ppid,name from category where ppid=3  and ClassType=1 order by id" 
rsClass3.open sqlClass3,cn,1,1
%>
var subval3 = new Array();
//����ṹ��������ֵ,������ֵ,������ʾֵ
<%
count3 = 0
do while not rsClass3.eof
%>
subval3[<%=count3%>] = new Array('<%=rsClass3("pID")%>','<%=rsClass3("ID")%>','<%=rsClass3("Name")%>')
<%
count3 = count3 + 1
rsClass3.movenext
loop
rsClass3.close
%>

function changeselect1(locationid)
{
    document.form1.pid.length = 0;
    document.form1.pid.options[0] = new Option('ѡ���������','');
    document.form1.ppid.length = 0;
    document.form1.ppid.options[0] = new Option('ѡ����������','');
    for (i=0; i<subval2.length; i++)
    {
        if (subval2[i][0] == locationid)
        {document.form1.pid.options[document.form1.pid.length] = new Option(subval2[i][2],subval2[i][1]);}
    }
}

function changeselect2(locationid)
{
    document.form1.ppid.length = 0;
    document.form1.ppid.options[0] = new Option('ѡ����������','');
    for (i=0; i<subval3.length; i++)
    {
        if (subval3[i][0] == locationid)
        {document.form1.ppid.options[document.form1.ppid.length] = new Option(subval3[i][2],subval3[i][1]);}
    }
}
//-->
</script><!-- ���������˵� ���� -->
	<%
Call header()

%>

  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.a_title.value == '' ) {
window.alert('���������±���^_^');
document.form1.a_title.focus();
return false;}

if ( document.form1.cid.value == '' ) {
window.alert('��ѡ�����^_^');
document.form1.cid.focus();
return false;}
return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>��������</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1����������ת��ַ������˼������������￴����һƪ�����£������ý����ݸ��Ƶ�������վ�У������ڴ���������µľ��Ե�ַ����ô�����ƪ���½���ֱ�������˵�ַ��</p>
 <p>2����ϵͳ�����ֶ���ҳ�����������Ҫ�����½��з�ҳ��ֻ���ڱ༭����ť�����<img src="images/inserthorizontalrule.gif" width="20" height="20" />ͼ�꣬�ͻ��ڱ༭�����Զ�����һ������Ϊ��ҳ��־���㲻��Ҫɾ������������������������ʾ��</p>  
            </td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRow'>���±��� (����) </td>
	<td class='forumRow'><input name='a_title' type='text' id='a_title' size='70'>
	  &nbsp;</td>
	</tr>
	<tr>
	<td class='forumRowHighLight' height=23>���·���<span class="forumRow"> (��ѡ) </span></td>
    <td class='forumRowHighLight'><%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select id,pid,ppid,name from category where ppid=1 and ClassType=1  order by id" 
rsClass1.open sqlClass1,cn,1,1
%>
            <select name="cid" id="cid" onChange="changeselect1(this.value)">
              <option>ѡ��һ������</option>
              <%
count1 = 0
do while not rsClass1.eof
response.write"<option value="&rsClass1("ID")&">"&rsClass1("Name")&"</option>"
count1 = count1 + 1
rsClass1.movenext
loop
rsClass1.close
%>
            </select>
            &nbsp;&nbsp;
            <select name="pid" id="pid"  onchange="changeselect2(this.value)">
              <option value="">ѡ���������</option>
            </select>
            &nbsp;&nbsp;
            <select name="ppid" id="ppid">
              <option value="">ѡ����������</option>
            </select>&nbsp;</td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>������ת��ַ</td>
	    <td class='forumRow'><input name='a_url' type='text' id='a_url' size='70'>
        &nbsp;��http://��ͷ</td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>����ͼƬ</td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%" ><input name="web_image" type="text" id="web_image"  size="30"></td>
           <td width="78%"  ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr>

        <td  class='forumRow' height=23>���¹ؼ���</td>
	      <td class='forumRow'><input type='text' id='a_keywords' name='a_keywords' size='60'> <select name="keywords_list" id="keywords_list" onclick="document.form1.a_keywords.value=document.form1.keywords_list.value">
	      <option value="">��ѡ��</option>
		   <% set rsp=server.createobject("adodb.recordset")
		   sql="select name from web_keywords order by [id] "
		   rsp.open(sql),cn,1,1
		   if not rsp.eof and not rsp.bof then
		   do while not rsp.eof 
		   %> <option value="<%=rsp("name")%>"  ><%=rsp("name")%></option>
            <%
			rsp.movenext
			loop
			end if
			rsp.close
			set rsp=nothing%></select>
	  &nbsp;���ԣ�����(���Ķ���)</td>
	</tr><tr>
	  <td class='forumRowHighLight' height=11>�������� / ����ժҪ </td>
	  <td class='forumRowHighLight'><textarea name='a_description'  cols="100" rows="4" id="a_description" ></textarea></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>�������� (����) </td>
	  <td class='forumRow'> <textarea name="a_content" id="a_content" style=" width:100%; height:400px; visibility:hidden;" ></textarea></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>������Դ</td>
	  <td class='forumRowHighLight'>
	    <input name='a_from_name' type='text' id='a_from_name' size='30'>
	 	      <select name="position" id="position" onclick="document.form1.a_from_name.value=document.form1.position.value">
	      <option value="">��ѡ��</option>
		   <% set rsp=server.createobject("adodb.recordset")
		   sql="select name from web_article_author order by [order] "
		   rsp.open(sql),cn,1,1
		   if not rsp.eof and not rsp.bof then
		   do while not rsp.eof 
		   %> <option value="<%=rsp("name")%>"  ><%=rsp("name")%></option>
            <%
			rsp.movenext
			loop
			end if
			rsp.close
			set rsp=nothing%></select>	 </td>
	  </tr>
	<tr>
	  <td class='forumRow' height=23>��Դ��ַ</td>
	  <td class='forumRow'><span class="forumRow">
	    <input name='a_from_url' type='text' id='a_from_url' size='40'>
      &nbsp;��http://��ͷ</span></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>��������</td>
	  <td class='forumRowHighLight'><span class="forumRow">
	    <input name='a_author' type='text' id='a_author' value="<%=Session("log_name")%>" size='40'>
	  </span></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>�����������</td>
	  <td class='forumRow'><input name='a_hit' type='text' id='a_hit' value="0" size='40'>
      &nbsp;ֻ��������</td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=23>�����ö�</td>
	  <td class='forumRowHighLight'><input type="radio" name="a_index_push" value="1">
��
      &nbsp;
      <input name="a_index_push" type="radio" value="0" checked>
��</td>
	  </tr>
	<tr>
	  <td class='forumRow' height=23>�滻�ؼ���</td>
	  <td class='forumRow'><input type="radio" name="a_keywords_yes" value="1">
��
      &nbsp;
      <input name="a_keywords_yes" type="radio" value="0" checked="checked" >
��</td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=23>��������</td>
	  <td class='forumRowHighLight'><label>
	    <input type="checkbox" name="slide_yes" value="1">
      ��Ϊ�õ�Ƭ 
      <input type="checkbox" name="special_yes" value="1">
      ��Ϊ�ȵ�ר��</label></td>
	  </tr>	  
	  
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
 %>