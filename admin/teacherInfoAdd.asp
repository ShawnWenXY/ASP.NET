<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/md5.asp"--> 
<%
  'errMessage���������Ϣ
  dim errMessage
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'�������Ա������µ�ѧ����Ϣ���ύ
	if Request("submit") <> "" then
	  '�����ʦְ�����û������
	  if Request("teacherNumber") = "" then
	    errMessage = "�������ʦ��ְ�����!"
	  end if
	  '�����ʦ�ĵ�½����û������
	  if Request("teacherPassword") = "" then
	    errMessage = "�������ʦ�ĵ�½����!"
		end if
	  if errMessage <> "" then
	    Response.Write "<script>alert('" & errMessage & "');</script>"
		else
	    '����ʦ������Ϣ��������ݿ���
		  set teacherInfoRs = Server.CreateObject("ADODB.RecordSet")
	    sqlString = "select * from [teacherInfo]"
	    teacherInfoRs.Open sqlString,conn,1,3
	    teacherInfoRs.AddNew
		  teacherInfoRs("teacherNumber") = Trim(Request("teacherNumber"))
		  teacherInfoRs("teacherName") = Trim(Request("teacherName"))
		  teacherInfoRs("teacherPassword") = md5(Trim(Request("teacherPassword")))
		  teacherInfoRs("teacherSex") = Trim(Request("teacherSex"))
		  '����ϴ��˽�ʦ��ͼƬ
		  if Request("photoAddress") <> "" then
		    teacherInfoRs("teacherPhoto") = Trim(Request("teacherPhoto"))
		  end if
		  '���ѡ���˽�ʦ������
		  if Request("teacherBirday") <> "" then
		    teacherInfoRs("teacherBirthday") = CDate(Request("teacherBirthday"))
		  else
		    teacherInfoRs("teacherBirthday") = CDate("1900-1-1")
			end if
		  '���ѡ���˽�ʦ����Уʱ��
		  if Request("teacherArriveTime") <> "" then
		    teacherInfoRs("teacherArriveTime") = CDate(Request("teacherArriveTime"))
		  else
		    teacherInfoRs("teacherArriveTime") = CDate("1900-1-1")
			end if
			teacherInfoRs("teacherCardNumber") = Trim(Request("teacherCardNumber"))
			teacherInfoRs("teacherAddress") = Trim(Request("teacherAddress"))
			teacherInfoRs("teacherPhone") = Trim(Request("teacherPhone"))
			teacherInfoRs("teacherMemo") = Trim(Request("teacherMemo"))
			teacherInfoRs.Update
			teacherInfoRs.Close
			Response.Write "<script>alert('��ʦ��Ϣ��ӳɹ�!');</script>"
	  end if
	end if
%>

<HTML>
<HEAD>
	<Title>�½�ʦ��Ϣ���</Title>
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
		      <img src="../images/ADD.gif" width=14px height=14px>��ʦ��Ϣ����--&gt;��ʦ��Ϣ���
			 </td>
	   </tr><br>
	   <tr>
	     <td style="height: 26px">
		     ��ʦְ�����:</td><td><input type=text name=teacherNumber size=20></td>
			 </td>
		 </tr>
		 <tr>
		  <td>��ʦ����:</td><td><input type=text name=teacherName size=20></td>
		 </tr>
		 <tr>
		   <td>�Ա�:</td>
			 <td>
			   <select name=teacherSex>
				   <option value='��'>��</option>
					 <option value='Ů'>Ů</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>��ʦ����:</td>
			 <td>
			   <input type=text name=teacherBirthday width=60px>
				 <input class="submit" name="Button" onclick="seltime('teacherBirthday');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
		   <td>��Уʱ��:</td>
			 <td>
			   <input type=text name=teacherArriveTime width=60px>
				 <input class="submit" name="Button" onclick="seltime('teacherArriveTime');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
		   <td>��½����:</td>
			 <td><input type=text name=teacherPassword size=20></td>
		 </tr>
		 <tr>
		   <td>��ʦ�绰:</td>
			 <td><input type=text name=teacherPhone size=20></td>
		 </tr>
		 <tr>
		    <td>���֤��:</td>
			  <td><input type=text name=teacherCardNumber size=40></td>
		  </tr>
		 <tr>
		    <td>��ͥ��ַ:</td>
			  <td><input type=text name=teacherAddress size=50></td>
		  </tr>
		 <tr>
			  <td>��Ƭ·��:</td>
			  <td><input type="text" name=photoAddress size=20 readonly="true">*���������ϴ���Ƭ,������Զ�����·��</td>
			</tr>
			<tr> 
       <td>��Ƭ�ϴ���</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=4 name=studentMemo></textarea></td>
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
