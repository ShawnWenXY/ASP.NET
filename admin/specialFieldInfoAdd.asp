<!--#include virtual="/Database/conn.asp"-->
<%
  'errMessage���������Ϣ
  dim errMessage,sqlString
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'�������Ա������µ�רҵ��Ϣ���ύ
	if Request("submit") <> "" then
	  '���û��ѡ�����ڵ�ѧԺ
	  if Request("specialCollegeNumber") = "" then
	    errMessage = "��ѡ������ѧԺ"
	  end if
	  '���û����дרҵ���
	  if Request("specialFieldNumber") = "" then
	    errMessage = "����дרҵ���"
	  end if
	  '���û����дרҵ����
	  if Request("specialFieldName") = "" then
	    errMessage = "����дרҵ������Ϣ"
	  end if
	  '���ϵͳ���Ƿ��Ѿ������˸�רҵ������Ϣ
	  sqlString = "select * from [specialFieldInfo] where specialFieldName='" & Request("specialFieldName") & "'"
	  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
	  specialFieldInfoRs.Open sqlString,conn,1,1
	  if not specialFieldInfoRs.EOF then
	    errMessage = "��רҵ������Ϣ�Ѿ����ڣ�"
		end if
	  specialFieldInfoRs.Close
	  '���errMessage��ֵ�Ƿ�Ϊ�վ����Ƿ�ִ����רҵ��Ϣ����Ӳ���
	  if errMessage = "" then
	    sqlString = "select * from [specialFieldInfo]"
		  specialFieldInfoRs.Open sqlString,conn,1,3
		  specialFieldInfoRs.AddNew
		  specialFieldInfoRs("specialCollegeNumber") = Request("specialCollegeNumber")
		  specialFieldInfoRs("specialFieldNumber") = Request("specialFieldNumber")
		  specialFieldInfoRs("specialFieldName") = Request("specialFieldName")
		  specialFieldInfoRs.Update
		  specialFieldInfoRs.Close
		  Response.Write "<script>alert('רҵ��Ϣ��ӳɹ�!');</script>"
	  else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
%>
<HTML>
<HEAD>
	<Title>רҵ��Ϣ���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
		<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	</script>
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr>
	     <td class="th" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�༶��Ϣ����--&gt;רҵ��Ϣ���
			 </td>
	   </tr><br>
		 <tr>
		  <td width=100 align="right">ѡ��ѧԺ:
			 <td>
			   <select name=specialCollegeNumber>
				   <option value="">��ѡ��</option>
				  <%
					  sqlString = "select collegeNumber,collegeName from [collegeInfo]"
					  set collegeInfoRs = Server.CreateObject("ADODB.RecordSet")
					  collegeInfoRs.Open sqlString,conn,1,1
					  while not collegeInfoRs.EOF
					    Response.Write "<option value='" & collegeInfoRs("collegeNumber") & "'>" & collegeInfoRs("collegeName") & "</option>"
						  collegeInfoRs.MoveNext
					  wend
				  %>
				 </select>
			 </td>
		 </tr>
		 <tr>
		    <td width=100 align="right">רҵ���:</td>
		    <td><input type=text name=specialFieldNumber size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">רҵ����:</td>
		    <td><input type=text name=specialFieldName size=20></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" ȷ����� ">
		      <input type="reset" value=" ������д ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>

</HTML>
