<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/album_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/album_content_to_html.asp" -->

<% '�������ݵ����ݱ�
act=Request("act")
If act="save" Then 
l_name=trim(request.form("name"))
l_url=trim(request.form("url"))
l_position=trim(request.form("position"))
l_image=trim(request.form("web_image"))
l_order=trim(request.form("order"))
l_memo=trim(request.form("memo"))
l_view_yes=trim(request.form("view_yes"))
a_index_push=trim(request.form("a_index_push"))
l_time=now()
picmore=trim(request.form("picmore"))

if picmore<>"" then


'��������ͼƬ
picmore1=split(picmore,",")
c=ubound(picmore1)
for i=0 to c
set rs=server.createobject("adodb.recordset")
sql="select * from web_ad"
rs.open(sql),cn,1,3
rs.addnew
rs("name")=l_name
rs("url")=l_url
rs("position")=l_position
If IsObjInstalled("Persits.Jpeg")  Then
rs("SmallImage")="small/"&picmore1(i)
end if
rs("image")=picmore1(i)
if l_order<>"" then
rs("order")=cint(l_order)
end if
rs("memo")=l_memo
rs("view_yes")=cint(l_view_yes)
'rs("index_push")=a_index_push

rs("time")=now()

rs.update
rs.close
set rs=nothing
next

end if

'�������չʾҳ
set rs_create=server.createobject("adodb.recordset")
sql="select [id],[name],[memo],backmusic from web_ad_position where [id]="&l_position
rs_create.open(sql),cn,1,1
if not rs_create.eof then
a_id=rs_create("id")
a_name=rs_create("name")
a_memo=rs_create("memo")
a_music=rs_create("backmusic")
end if
rs_create.close
set rs_create=nothing
call album_content_to_html(a_id,a_name,a_memo,a_music)
call album_index_to_html()
call index_to_html()
response.Write "<script language='javascript'>alert('���ӳɹ���');location.href='ad_list.asp';</script>"
end if 
 %>
	<%
Call header()

%>
 <script type="text/javascript" src="PicUpload/init.js"></script>

  <form id="form1" name="form1" method="post" action="?act=save"  >
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('������ͼƬ����^_^');
document.form1.name.focus();
return false;}


if ( document.form1.position.value == '' ) {
window.alert('��ѡ�����^_^');
document.form1.position.focus();
return false;}

if ( document.form1.picmore.value == '' ) {
window.alert('���ϴ�����һ��ͼƬ^_^');
document.form1.picmore.focus();
return false;}

if(document.form1.order.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("����ֻ��������^_^");   
document.form1.order.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>�����ϴ�ͼƬ</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>
<p>1��ͼƬֻ֧��jpg,gif,bmp��ʽ����С���ܳ���1024K(1.0M)��</p>
            <p>2��ͨ��������������ͼƬһ�㶼�Ƚϴ󣬴ＸM����Ҫͨ��PHOTOSHOP���д�����</p>
            <p>3���ϴ���ͼƬ�����ڸ�Ŀ¼�µ�photos�ļ����ڡ�</p>
			<p>4��ͼƬ�޷��ϴ�����������ԭ��(1)��Ŀռ䲻֧��FSO�����(2)��Ŀռ�д��Ȩ��δ�򿪣�(3)��ʹ�õ��Ǽ���ASP����������ѿռ���в��ԣ�(4)���ͼƬ��ʽ���Ի��ļ��������ƣ�(5)ͼƬ����ļ��в����ڣ�(6)��Ŀռ�������(7)��Ŀռ��ٶȹ��ͣ�(8)�ڿ������ˡ�</p>
			<p>5�������ȷ�����������û�г��ֵĻ�����ô������ϵ���������ˡ�</p>			</td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>ͼƬ���� (����)</td>
	<td width="85%" class='forumRowHighLight'><input name='name' type='text' id='name' size='70'>
	  &nbsp;</td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>ͼƬ���� </td>
	    <td class='forumRow'><input name='url' type='text' id='url' value="" size='70'>
        &nbsp;��http://��ͷ</td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>ѡ�����<span class="forumRowHighLight"> (��ѡ)</span></td>
	    <td class='forumRowHighLight'><label>
	      <select name="position" id="position">
	       <% set rsp=server.createobject("adodb.recordset")
		   sql="select id,name from web_ad_position "
		   rsp.open(sql),cn,1,1
		   if not rsp.eof and not rsp.bof then
		   do while not rsp.eof 
		   %> <option value="<%=rsp("id")%>"><%=rsp("name")%></option>
            <%
			rsp.movenext
			loop
			end if
			rsp.close
			set rsp=nothing%></select>
	    </label></td>
      </tr>
<tr>
	    <td class='forumRow' height=23>�ϴ�ͼƬ</td>
	    <td class='forumRow'><input id="picmore" name="picmore" type="text" size="80" /> <br><input type="button" value="�ϴ�ͼƬ" onClick="showUpload(null,'picmore','',999,null);" /></td>
      </tr>	  <tr>
	    <td class='forumRowHighLight' height=23>����</td>
	    <td class='forumRowHighLight'><span class="forumRow">
	      <input name='order' type='text' id='order' value="100" size='20'>
	    &nbsp;ֻ�������֣�����ԽС����Խ��ǰ</span></td>
      </tr>
	  <tr>
	  <td class='forumRow' height=11>����</td>
	  <td class='forumRow'><textarea name='memo'  cols="100" rows="6" id="memo" ></textarea></td>
	</tr>
	  
	  <tr>
	  <td class='forumRowHighLight' height=23>�Ƿ���ʾ</td>
	  <td class='forumRowHighLight'><label>
	    <input type="radio" name="view_yes" value="1" checked>
      ��
      &nbsp;
      <input name="view_yes" type="radio" value="0" >
      ��</label></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
Call DbconnEnd()
 %>