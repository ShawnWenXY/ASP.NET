<!--#include virtual="/DataBase/conn.asp"-->
<%
  dim sqlString,systemName,author,email,qq,phone,pageSize,canSelect,termId,selectStartTime,selectEndTime
   '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'����޸���������ϢҪ�����
	if Request("submit") <> "" then
	  sqlString = "select * from [config] where id=1"
	  set configInfoRs = Server.CreateObject("ADODB.RecordSet")
	  configInfoRs.Open sqlString,conn,1,3
	  configInfoRs("systemName") = Trim(Request("systemName"))
	  configInfoRs("author") = Trim(Request("author"))
	  configInfoRs("email") = Trim(Request("email"))
	  configInfoRs("qq") = Trim(Request("qq"))
	  configInfoRs("phone") = Trim(Request("phone"))
	  configInfoRs("pageSize") = CInt(Request("pageSize"))
	  configInfoRs("canSelect") = CInt(Request("canSelect"))
	  configInfoRs("termId") = CInt(Request("termId"))
	  configInfoRs("selectStartTime") = CDate(Request("selectStartTime"))
	  configInfoRs("selectEndTime") = CDate(Request("selectEndTime"))
	  configInfoRs.Update
	  Response.Write "<script>alert('ϵͳ������Ϣ���³ɹ�!  ');</script>"
	end if
	sqlString = "select * from [config] where id=1"
	set configInfoRs = Server.CreateObject("ADODB.RecordSet")
	configInfoRs.Open sqlString,conn,1,1
	systemName = configInfoRs("systemName")
	author = configInfoRs("author")
	email = configInfoRs("email")
	qq = configInfoRs("qq")
	phone = configInfoRs("phone")
	pageSize = CInt(configInfoRs("pageSize"))
	canSelect = CInt(configInfoRs("canSelect"))
	termId = CInt(configInfoRs("termId"))
	selectStartTime = CDate(configInfoRs("selectStartTime"))
	selectEndTime = CDate(configInfoRs("selectEndTime"))
%>
<HTML>
<HEAD>
	<Title>ϵͳ��������</Title>
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
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>ϵͳ��������
			 </td>
	   </tr><br>
		<tr>
		  <td width=100 align="right">ϵͳ����:</td>
		  <td><input type=text name=systemName size=30 value='<%=systemName%>'></td>
		</tr>
		<tr>
		  <td width=100 align="right">��������:</td>
		  <td><input type=text name=author size=20 value='<%=author%>'></td>
		</tr>
		<tr>
		  <td width=100 align="right">Email:</td>
		  <td><input type=text name=email size=30 value='<%=email%>'></td>
		</tr>
		<tr>
		  <td width=100 align="right">QQ:</td>
		  <td><input type=text name=qq size=20 value='<%=qq%>'></td>
		</tr>
		<tr>
		  <td width=100 align="right">�绰:</td>
		  <td><input type=text name=phone size=20 value='<%=phone%>'></td>
		</tr>
		<tr>
		  <td width=100 align="right">ÿҳ��Ϣ����:</td>
		  <td><input type=text name=pageSize size=5 value='<%=pageSize%>'>��</td>
		</tr>
		<tr>
		  <td width=100 align="right">�Ƿ񿪷�ѡ��:</td>
		  <td>
		    <select name=canSelect>
		    <%
			     if canSelect = 1 then
				     Response.Write "<option value='1'>��</option><option value='0'>��</option>"
					 else
					   Response.Write "<option value='0'>��</option><option value='1'>��</option>"
					 end if
			  %>
			  </select>
		  </td>
		</tr>
		<tr>
		  <td width=100 align="right">����ѡ��ѧ��:</td>
		  <td>
		    <select name=termId>
		    <%
			    sqlString = "select * from [termInfo]"
				  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
				  termInfoRs.Open sqlString,conn,1,1
				  while not termInfoRs.EOF
				    selected = ""
					  if termInfoRs("termId") = configInfoRs("termId") then
					    selected = "selected"
					  end if
					  Response.Write "<option value='" & termInfoRs("termId") & "' " & selected & ">" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "���" & termInfoRs("termUpOrDown") & "</option>"
					  termInfoRs.MoveNext
 				  wend
			  %>
			  </select>
		  </td>
		</tr>
		<tr>
		  <td width=100 align="right">ѡ�ο�ʼ����:</td>
		  <td>
		    <input type=text name=selectStartTime width=77px value='<%=configInfoRs("selectStartTime")%>'>
				<input class="submit" name="Button" onclick="seltime('selectStartTime');" style="width:30px" type="button" value="ѡ��">
			</td>
		</tr>
		<tr>
		  <td width=100 align="right">ѡ�ν�������:</td>
		  <td>
		    <input type=text name=selectEndTime width=77px value='<%=configInfoRs("selectEndTime")%>'>
				<input class="submit" name="Button" onclick="seltime('selectEndTime');" style="width:30px" type="button" value="ѡ��">
			</td>
		</tr>
		<tr>
		  <td colspan=2 align="center"><input type="submit"  name=submit value="����"></td>
		</tr>
	</table>
</form>
</body>
</html>