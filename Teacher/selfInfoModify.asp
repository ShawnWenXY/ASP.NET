<!--#include file="../Database/conn.asp"-->
<%
  '�����ʦ��û�е�½
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'��������˽�ʦ�������Ϣ���ύʱ
	if Request("submit") <> "" then
	  sqlString = "select * from [teacherInfo] where teacherNumber='" & Session("teacherNumber") & "'"
	  set teacherInfoRs = Server.CreateObject("ADODB.RecordSet")
	  teacherInfoRs.Open sqlString,conn,1,3
	  teacherInfoRs("teacherName") = Request("teacherName")
	  teacherInfoRs("teacherSex") = Request("teacherSex")
	  '�������ѡ���˽�ʦ�ĸ���ͷ��
	  if Request("photoAddress") <> "" then
	    teacherInfoRs("teacherPhoto") = Trim(Request("photoAddress"))
	  end if
	  teacherInfoRs("teacherBirthday") = CDate(Request("teacherBirthday"))
	  teacherInfoRs("teacherArriveTime") = CDate(Request("teacherArriveTime"))
	  teacherInfoRs("teacherCardNumber") = Trim(Request("teacherCardNumber"))
	  teacherInfoRs("teacherAddress") = Trim(Request("teacherAddress"))
	  teacherInfoRs("teacherPhone") = Trim(Request("teacherPhone"))
	  teacherInfoRs("teacherMemo") = Trim(Request("teacherMemo"))
	  teacherInfoRs.Update
	  teacherInfoRs.Close
	  Response.Write "<script>alert('��ʦ��Ϣ���³ɹ�!');</script>"
	end if
  '�õ�ĳ����ʦ����ϸ��Ϣ
  set teacherDetailRs = Server.CreateObject("ADODB.RecordSet")
  sqlString = "select * from [teacherInfo] where teacherNumber='" & Session("teacherNumber") & "'"
  teacherDetailRs.Open sqlString,conn,1,1
%>
<HTML>
<HEAD>
	<Title>������Ϣ�޸�</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript">
	function seltime(inputName)
	{
	  window.open('../admin/seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	</script>
</HEAD>
<BODY>
<form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>ϵͳ����--&gt;������Ϣ�޸�
			 </td>
	   </tr><br>
		 <%
		   '����ý�ʦ������ͼƬ����ʾ�ý�ʦ��ͷ��
		   if teacherDetailRs("teacherPhoto") <> "" then
			   Response.Write "<tr><td>��ʦͷ��:</td><td><img src='../admin/" & teacherDetailRs("teacherPhoto") & "' border=0 height=100 width=100></td></tr>"
			 end if 
		 %>
	   <tr>
	     <td style="height: 26px">
		     ��ʦְ�����:</td><td><%=teacherDetailRs("teacherNumber")%></td>
			 </td>
		 </tr>
		 <tr>
		  <td>��ʦ����:</td><td><input type=text name=teacherName size=20 value=<%=teacherDetailRs("teacherName")%>></td>
		 </tr>
		 <tr>
		   <td>�Ա�:</td>
			 <td>
			   <select name=teacherSex>
			   <%
				   if teacherDetailRs("teacherSex") = "��" then
					   Response.Write "<option value='��'>��</option><option value='Ů'>Ů</option>"
					 else
					   Response.Write "<option value='Ů'>Ů</option><option value='��'>��</option>"
					 end if
				 %>
			 </td>
		 </tr>
		 <tr>
		   <td>��ʦ����:</td>
			 <td>
			   <input type=text name=teacherBirthday width=77px value=<%=teacherDetailRs("teacherBirthday")%>>
				 <input class="submit" name="Button" onclick="seltime('teacherBirthday');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
		   <td>��Уʱ��:</td>
			 <td>
			   <input type=text name=teacherArriveTime width=77px value=<%=teacherDetailRs("teacherArriveTime")%>>
				 <input class="submit" name="Button" onclick="seltime('teacherArriveTime');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
			  <td>����Ƭ·��:</td>
			  <td><input type="text" name=photoAddress size=20 readonly>*���������ϴ���Ƭ,������Զ�����·��</td>
			</tr>
			<tr> 
       <td>����Ƭ�ϴ���</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='../admin/upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  <tr>
		    <td>���֤��:</td>
			  <td><input type=text name=teacherCardNumber size=50 value=<%=teacherDetailRs("teacherCardNumber")%>></td>
		  </tr>
		  <tr>
		    <td>��ͥ��ַ:</td>
			  <td><input type=text name=teacherAddress size=50 value=<%=teacherDetailRs("teacherAddress")%>></td>
		  </tr>
		  <tr>
		    <td>�绰:</td>
			  <td><input type=text name=teacherPhone size=50 value=<%=teacherDetailRs("teacherPhone")%>></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=teacherMemo><%=teacherDetailRs("teacherMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" ȷ�ϸ��� ">
		      <input type="button" value=" ����" onClick="javascript:location.href='teacherInfoManage.asp';">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
