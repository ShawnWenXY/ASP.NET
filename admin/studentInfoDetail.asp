<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/md5.asp"--> 
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'���������ѧ���������Ϣ���ύʱ
	if Request("submit") <> "" then
	  sqlString = "select * from [studentInfo] where studentNumber='" & Request("studentNumber") & "'"
	  set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
	  studentInfoRs.Open sqlString,conn,1,3
	  studentInfoRs("studentName") = Request("studentName")
	  studentInfoRs("studentSex") = Request("studentSex")
	  studentInfoRs("studentBirthday") = CDate(Request("studentBirthday"))
	  studentInfoRs("studentState") = Request("studentState")
	  '��������ѧ�������µĵ�½����
	  if Request("studentPassword") <> "" then
	    studentInfoRs("studentPassword") = md5(Request("studentPassword"))
	  end if
	  '�������Ա��ѧ���ϴ����µ�ͼƬ 
	  if Request("photoAddress") <> "" then
	    studentInfoRs("studentPhoto") = Request("photoAddress")
		end if
	  studentInfoRs("studentAddress") = Request("studentAddress")
	  studentInfoRs("studentMemo") = Request("studentMemo")
	  studentInfoRs.Update
	  studentInfoRs.Close
	  Response.Write "<script>alert('ѧ����Ϣ���³ɹ�!');</script>"
	end if
  '�õ�ĳ��ѧ������ϸ��Ϣ
  set studentDetailRs = Server.CreateObject("ADODB.RecordSet")
  sqlString = "select * from [studentInfo] where studentNumber='" & Request("studentNumber") & "'"
  studentDetailRs.Open sqlString,conn,1,1
%>
<HTML>
<HEAD>
	<Title>ѧ����ϸ��Ϣ</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript">
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	</script>
</HEAD>
<BODY>
<form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>ѧ����Ϣ����--&gt;ѧ����ϸ��Ϣ
			 </td>
	   </tr><br>
		 <%
		   '�����ѧ��������ͼƬ����ʾ��ѧ����ͷ��
		   if studentDetailRs("studentPhoto") <> "" then
			   Response.Write "<tr><td>ѧ��ͷ��:</td><td><img src='" & studentDetailRs("studentPhoto") & "' border=0 height=100 width=100></td></tr>"
			 end if 
		 %>
		 <tr>
			 <td>���ڰ༶:</td>
			 <td>
			   <select name=studentClass id=studentClass>
				  <%
				    '�õ����еİ༶��Ϣ
				    set studentClassRs = Server.CreateObject("ADODB.RecordSet")
					  sqlString = "select classNumber,className from [classInfo]"
					  studentClassRs.Open sqlString,conn,1,1
					  while not studentClassRs.EOF
					    selected = ""
						  if studentClassRs("classNumber") = studentDetailRs("studentClassNumber") then
						    selected = "selected"
							end if
					    Response.Write "<option value='" & studentClassRs("classNumber") &"' " & selected & ">" & studentClassRs("className") & "</option>"
						  studentClassRs.MoveNext
					  wend
				  %>
				</select>
			 </td>
		 </tr>
	   <tr>
	     <td style="height: 26px">
		     ѧ��:</td><td><%=studentDetailRs("studentNumber")%></td>
			   <input type="hidden" name=studentNumber value=<%=studentDetailRs("studentNumber")%>>
			 </td>
		 </tr>
		 <tr>
		  <td>ѧ������:</td><td><input type=text name=studentName size=20 value=<%=studentDetailRs("studentName")%>></td>
		 </tr>
		 <tr>
		   <td>�Ա�:</td>
			 <td>
			   <select name=studentSex>
			   <%
				   if studentDetailRs("studentSex") = "��" then
					   Response.Write "<option value='��'>��</option><option value='Ů'>Ů</option>"
					 else
					   Response.Write "<option value='Ů'>Ů</option><option value='��'>��</option>"
					 end if
				 %>
			 </td>
		 </tr>
		 <tr>
		   <td>ѧ������:</td>
			 <td>
			   <input type=text name=studentBirthday width=77px value=<%=studentDetailRs("studentBirthday")%>>
				 <input class="submit" name="Button" onclick="seltime('studentBirthday');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
		   <td>������ò:</td>
			 <td>
			   <select name="studentState">
				   <option value='��Ա' <% if studentDetailRs("studentState")="��Ա" then Response.Write "selected" end if%>>��Ա</option>
					 <option value='��Ա' <% if studentDetailRs("studentState")="��Ա" then Response.Write "selected" end if%>>��Ա</option>
					 <option value='�ϰ���' <% if studentDetailRs("studentState")="�ϰ���" then Response.Write "selected" end if%>>�ϰ���</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>��½����:</td>
			 <td><input type=text name=studentPassword size=20><font color=red>���ҪΪ��ѧ�����������������ڴ�����</font></td>
		 </tr>
		 
		 <tr>
			  <td>����Ƭ·��:</td>
			  <td><input type="text" name=photoAddress size=20 readonly="true">*���������ϴ���Ƭ,������Զ�����·��</td>
			</tr>
			<tr> 
       <td>����Ƭ�ϴ���</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  <tr>
		    <td>��ͥ��ַ:</td>
			  <td><input type=text name=studentAddress size=50 value=<%=studentDetailRs("studentAddress")%>></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=studentMemo><%=studentDetailRs("studentMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" ȷ�ϸ��� ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
