<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<%
act=Request("act")
If act="save" Then 
web_ListTime=trim(request.form("web_ListTime"))
web_ListIntro=trim(request.form("web_ListIntro"))
web_ListCount=trim(request.form("web_ListCount"))
web_ListKeywords=trim(request.form("web_ListKeywords"))
web_ListAuthor=trim(request.form("web_ListAuthor"))
web_FeedComment=trim(request.form("web_FeedComment"))
web_FeedAdvice=trim(request.form("web_FeedAdvice"))
web_FeedCount=trim(request.form("web_FeedCount"))
web_FeedTime=trim(request.form("web_FeedTime"))
web_SideImage=trim(request.form("web_SideImage"))
web_SideClass=trim(request.form("web_SideClass"))
web_SideHot=trim(request.form("web_SideHot"))
web_time=now()


set rs=server.createobject("adodb.recordset")
sql="select * from web_AdvancedSettings"
rs.open(sql),cn,1,3
rs("web_ListTime")=web_ListTime
rs("web_ListIntro")=web_ListIntro
rs("web_ListCount")=web_ListCount
rs("web_ListKeywords")=web_ListKeywords
rs("web_ListAuthor")=web_ListAuthor
rs("web_FeedComment")=web_FeedComment
rs("web_FeedAdvice")=web_FeedAdvice
rs("web_FeedCount")=web_FeedCount
rs("web_FeedTime")=web_FeedTime
rs("web_SideImage")=web_SideImage
rs("web_SideClass")=web_SideClass
rs("web_SideHot")=web_SideHot
rs("web_time")=web_time
rs.update
rs.close
set rs=nothing

call index_to_html()
response.Write "<script language='javascript'>alert('�޸ĳɹ���')</script>"

end if
 %>
<script type="text/javascript" charset="utf-8" src="../KKKeditor/kindeditor.js"></script>
<script type="text/javascript" src="../KKKeditor/editor.js"></script>	
	<%
Call header()

%>
<%set rs=server.createobject("adodb.recordset")
sql="select * from web_AdvancedSettings"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
%>
  <form id="form1" name="form1" method="post" action="?act=save">
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=31>��վ�߼�����</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>
            <p>1���޸���ĳ����Ϣ��Ĭ��ֻ���Զ�������վ��ҳ������ҳ����Ҫ�ֶ���"���ɹ���"��<a href="html_items.asp">������Ŀ</a>��<a href="html_article.asp">��������</a>�Żῴ���޸ĺ��Ч����</p></td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td height=30 colspan="2" class='forumRowHighLight'><strong>&nbsp;&nbsp;&nbsp;&nbsp;�б�����</strong></td>
	</tr>
	<tr>
	<td width="15%" height=23 class='forumRow'> ʱ��</td>
	<td class='forumRow'><label>
	       <input type="radio" name="web_ListTime" value="1"<%
		if rs("web_ListTime")=1 then
		response.write "checked"
		end if%>>
      ��ʾ
      &nbsp;
      <input name="web_ListTime" type="radio" value="0" <%if rs("web_ListTime")=0 then
		response.write "checked"
		end if%>>
      ����ʾ</label></td>
	</tr>    
	<tr>
	  <td class='forumRow' height=23> ���</td>
	  <td class='forumRow'><label>
	       <input type="radio" name="web_ListIntro" value="1"<%
		if rs("web_ListIntro")=1 then
		response.write "checked"
		end if%>>
      ��ʾ
      &nbsp;
      <input name="web_ListIntro" type="radio" value="0" <%if rs("web_ListIntro")=0 then
		response.write "checked"
		end if%>>
      ����ʾ</label></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRow'> �ؼ���</td>
	<td class='forumRow'><label>
	       <input type="radio" name="web_ListKeywords" value="1"<%
		if rs("web_ListKeywords")=1 then
		response.write "checked"
		end if%>>
      ��ʾ
      &nbsp;
      <input name="web_ListKeywords" type="radio" value="0" <%if rs("web_ListKeywords")=0 then
		response.write "checked"
		end if%>>
      ����ʾ</label></td>
	</tr>    
	<tr>
	  <td class='forumRow' height=23> ����</td>
	  <td class='forumRow'><label>
	       <input type="radio" name="web_ListAuthor" value="1"<%
		if rs("web_ListAuthor")=1 then
		response.write "checked"
		end if%>>
      ��ʾ
      &nbsp;
      <input name="web_ListAuthor" type="radio" value="0" <%if rs("web_ListAuthor")=0 then
		response.write "checked"
		end if%>>
      ����ʾ</label></td>
	  </tr>      
	<tr>
	<td width="15%" height=23 class='forumRow'> ��������</td>
	<td class='forumRow'><label>
          <input type="radio" name="web_ListCount" value="5"<%
		if rs("web_ListCount")=5 then
		response.write "checked"
		end if%>>
      5��
      &nbsp;
      <input name="web_ListCount" type="radio" value="10" <%if rs("web_ListCount")=10 then
		response.write "checked"
		end if%>>
      10��      &nbsp;
      <input name="web_ListCount" type="radio" value="20" <%if rs("web_ListCount")=20 then
		response.write "checked"
		end if%>>
      20��
</label></td>
	</tr>  
	<tr>
	<td height=30 colspan="2" class='forumRowHighLight'><strong>&nbsp;&nbsp;&nbsp;&nbsp;������������</strong></td>
	</tr>
	<tr>
	<td width="15%" height=23 class='forumRow'> ��������</td>
	<td class='forumRow'><label>
	       <input type="radio" name="web_FeedComment" value="1"<%
		if rs("web_FeedComment")=1 then
		response.write "checked"
		end if%>>
      ���
      &nbsp;
      <input name="web_FeedComment" type="radio" value="0" <%if rs("web_FeedComment")=0 then
		response.write "checked"
		end if%>>
      �����</label></td>
	</tr>    
	<tr>
	  <td class='forumRow' height=23> �ÿ�����</td>
	  <td class='forumRow'><label>
	       <input type="radio" name="web_FeedAdvice" value="1"<%
		if rs("web_FeedAdvice")=1 then
		response.write "checked"
		end if%>>
      ���
      &nbsp;
      <input name="web_FeedAdvice" type="radio" value="0" <%if rs("web_FeedAdvice")=0 then
		response.write "checked"
		end if%>>
      �����</label></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRow'> ��������</td>
	<td class='forumRow'><label>
	       <input type="radio" name="web_FeedCount" value="5"<%
		if rs("web_FeedCount")=5 then
		response.write "checked"
		end if%>>
      5��
      &nbsp;
      <input name="web_FeedCount" type="radio" value="10" <%if rs("web_FeedCount")=10 then
		response.write "checked"
		end if%>>
      10��      &nbsp;
      <input name="web_FeedCount" type="radio" value="20" <%if rs("web_FeedCount")=20 then
		response.write "checked"
		end if%>>
      20��
</label></td>
	</tr>  
 	<tr>
	  <td class='forumRow' height=23> ����</td>
	  <td class='forumRow'><label>
	       <input type="radio" name="web_FeedTime" value="desc"<%
		if rs("web_FeedTime")="desc" then
		response.write "checked"
		end if%>>
      ��ʱ������
      &nbsp;
      <input name="web_FeedTime" type="radio" value="asc" <%if rs("web_FeedTime")="asc" then
		response.write "checked"
		end if%>>
      ��ʱ�䵹��</label></td>
	  </tr>   
	<tr>
	<td height=30 colspan="2" class='forumRowHighLight'><strong>&nbsp;&nbsp;&nbsp;&nbsp;���������</strong></td>
	</tr>
	<tr>
	<td width="15%" height=23 class='forumRow'> ͼƬ��ʾ��</td>
	<td class='forumRow'><label>
	       <input type="radio" name="web_SideImage" value="4"<%
		if rs("web_SideImage")=4 then
		response.write "checked"
		end if%>>
      4��
      &nbsp;
      <input name="web_SideImage" type="radio" value="6" <%if rs("web_SideImage")=6 then
		response.write "checked"
		end if%>>
      6��      &nbsp;
      <input name="web_SideImage" type="radio" value="8" <%if rs("web_SideImage")=8 then
		response.write "checked"
		end if%>>
      8��
</label></td>
	</tr>    
	<tr>
	  <td class='forumRow' height=23> ������ʾ��</td>
	  <td class='forumRow'><label>
	       <input type="radio" name="web_SideClass" value="4"<%
		if rs("web_SideClass")=4 then
		response.write "checked"
		end if%>>
      4��
      &nbsp;
      <input name="web_SideClass" type="radio" value="6" <%if rs("web_SideClass")=6 then
		response.write "checked"
		end if%>>
      6��      &nbsp;
      <input name="web_SideClass" type="radio" value="8" <%if rs("web_SideClass")=8 then
		response.write "checked"
		end if%>>
      8��
</label></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRow'> ������ʾ��</td>
	<td class='forumRow'><label>
	       <input type="radio" name="web_SideHot" value="5"<%
		if rs("web_SideHot")=5 then
		response.write "checked"
		end if%>>
      5��
      &nbsp;
      <input name="web_SideHot" type="radio" value="10" <%if rs("web_SideHot")=10 then
		response.write "checked"
		end if%>>
      10��      &nbsp;
      <input name="web_SideHot" type="radio" value="15" <%if rs("web_SideHot")=15 then
		response.write "checked"
		end if%>>
      15��
</label></td>
	</tr>                
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
else
response.write "��ʱ������"
end if %>