<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->
<% '����ģ��
act=request.querystring("act")
keywords=trim(request.form("keywords"))
if act="search" then
if keywords<>"" then
s_sql="select * from web_Files where [FileName] like '%"&keywords&"%' order by [FileTime] desc"
else
s_sql="select * from web_Files where [FileName] like '%"&keywords&"%' order by [FileTime] desc"
end if
else
s_sql="select * from web_Files where [FileName] like '%"&keywords&"%' order by [FileTime] desc"
end if 
%>

<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='���棺ɾ���󽫲��ɻָ�������������벻�������';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>

<script type="text/javascript">function copyText(obj)   
{  
var rng = document.body.createTextRange();  
rng.moveToElementText(obj);  
rng.scrollIntoView();  
rng.select();  
rng.execCommand("Copy");  
rng.collapse(false);
alert("�ļ����ص�ַ���Ƴɹ�������Է���������ѻ����������վ��Ŷ^_^"); 
}  
</script>  
	<%
Call header()
%>
<%
set rs2=server.createobject("adodb.recordset")
sql="select * from web_FileSetting "
rs2.open(sql),cn,1,1
if not rs2.eof  then
FileFolder=rs2("FileFolder")
FileType=rs2("FileType")
FileSize=rs2("FileSize")
end if
rs2.close
set rs2=nothing
%>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>���ϴ��ļ��б�</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;�� ������ʾ</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1��Ŀǰ�����ļ��������ϵͳ��Ŀ¼�� <%=FileFolder%> �ļ����¡�</p>
              <p>2�����鶨��ɾ������Ҫ���ļ��Խ�ʡϵͳ�ռ䡣</p>
			  <p>3����������ļ����ص�ַ������ȷ�ϡ�<a href="web_settings.asp">��վϵͳ����</a>��- ����վ��ַ������������ȷ�ġ�</p>
			  <p>4��ɾ���ļ����޷��ָ��������������</p></td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
      </table>	
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="File_Upload.asp">�ϴ����ļ�</a></td>
          </tr>
          <tr>
            <td height="30"></td>
          </tr>
      </table>
	   
	  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="5%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���</div></td>
            <td width="21%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">�ļ���</div></td>
            <td width="28%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">��ע</div></td>
            <td width="17%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">�ϴ�ʱ��</div></td>
            <td width="14%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">�������ص�ַ</div></td>
            <td width="15%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">����</div></td>
          </tr>
<% '�����б�ģ��
strFileName="File_list.asp" 
pageno=20
set rs = server.CreateObject("adodb.recordset")

rs.Open (s_sql),cn,1,1
rscount=rs.recordcount
if not rs.eof and not rs.bof then
call showsql(pageno)
rs.move(rsno)
for p_i=1 to loopno
%>
<% if p_i mod 2 =0 then
class_style="forumRow"
else
class_style="forumRowHighLight"
end if%>
            <form name="form1" method="post" action="?action=edit&id=<%=rs("id")%>">
          <tr >
            <td   height="40" class='<%=class_style%>'><div align="center"><%=rs("id")%></div></td>
           <td class='<%=class_style%>' ><div align="center"><a href="<%="/"&FileFolder&"/"&rs("FileName")%>" target="_blank"><%=rs("FileName")%></a> [<font color="#FF0000"><%=rs("FileSize")%>KB</font>]</div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("FileMemo")%></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("FileTime")%></div></td>
            <td class='<%=class_style%>' ><div align="center"><span id="tbid<%=p_i%>"><%=WebUrl&FileFolder&"/"&rs("FileName")%></span> <a href="#" onClick="copyText(document.all.tbid<%=p_i%>)">�������</a></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="<%="/"&FileFolder&"/"&rs("FileName")%>" target="_blank" title="�����һ���ѡ��Ŀ�����Ϊ������">����</a> | <a href="File_Edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>" >�޸�</a> | <a href="javascript:if(ask('���棺ɾ���󽫲��ɻָ�������������벻�������')) location.href='File_del.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">ɾ��</a>            </div></td>
          </tr></form>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>�������ݣ�</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>
		    <tr  >
              <td height="35"  colspan="10" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| �ļ�����</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search"><div align="center">&nbsp;
              <label>
                </label>
            <label>
<input name="keywords" type="text"  size="35" maxlength="40">
              </label>
                <label>
                       &nbsp;
                       <input type="submit" name="Submit" value="�� ��">
                </label>
              </div>
            </form>
            </td>
          </tr>
      </table>	  
</td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>