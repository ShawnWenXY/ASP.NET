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
	  '���û��ѡ��༶
	  if Request("studentClass") ="" then
	    errMessage = "��ѡ��ѧ�����ڵİ༶!"
	  end if
	  '���ѧ��û������
	  if Request("studentNumber") = "" then
	    errMessage = "������ѧ����ѧ��!"
	  end if
	  '���ѧ���ĵ�½����û������
	  if Request("studentPassword") = "" then
	    errMessage = "������ѧ���ĵ�½����!"
		end if
	  if errMessage <> "" then
	    Response.Write "<script>alert('" & errMessage & "');</script>"
		else
	    '��ѧ��������Ϣ��������ݿ���
	    set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
	    sqlString = "select * from [studentInfo]"
	    studentInfoRs.Open sqlString,conn,1,3
	    studentInfoRs.AddNew
	    studentInfoRs("studentNumber") = Trim(Request("studentNumber"))
	    studentInfoRs("studentName") = Trim(Request("studentName"))
	    studentInfoRs("studentPassword") = md5(Trim(Request("studentPassword")))
	    studentInfoRs("studentSex") = Trim(Request("studentSex"))
	    studentInfoRs("studentState") = Trim(Request("studentState"))
	    studentInfoRs("studentPhoto") = Trim(Request("photoAddress"))
	    studentInfoRs("studentClassNumber") = Trim(Request("studentClass"))
	    '�������Աѡ����ѧ��������
	    if Request("studentBirthday") <> "" then
	      studentInfoRs("studentBirthday") = CDate(Request("studentBirthday"))
		  else
		    studentInfoRs("studentBirthday") = CDate("1900-1-1")
		  end if
		  studentInfoRs("studentAddress") = Request("studentAddress")
		  studentInfoRs("studentMemo") = Request("studentMemo")
		  studentInfoRs.Update
		  studentInfoRs.Close
		  Response.Write "<script>alert('ѧ����Ϣ��ӳɹ�!')</script>"
	  end if
	end if
%>

<HTML>
<HEAD>
	<Title>��ѧ����Ϣ���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	
	var college_specialField = new Array();
	var specialField_class = new Array();
	//��ʼ������(ѧԺ��רҵ)��Ϣ��¼����
  function initCSArray() {
   
    Server.ScriptTimeout = "10"
    set conn=server.CreateObject("Adodb.Connection")
    Path="driver={SQL Server};server=.;uid=sa;pwd=123456;database=SchoolManage" 
    conn.open path
    dim i  'ѭ������
	  i = 0
    sql = "select count(*) as count from [specialFieldInfo]"
	'set countRs = conn.Execute(sql)
   ' count = countRs("count") '�õ��ܵļ�¼��
    sql = "select * from specialFieldInfo" 
	  set specialFieldRs = conn.Execute(sql)
	  
	  while not specialFieldRs.eof
	    Response.Write "college_specialField[" & i & "]='" & specialFieldRs("specialCollegeNumber") & ":" &specialFieldRs("specialFieldNumber") & ":" &specialFieldRs("specialFieldName") & "';" & vbCrLf
		  i = i + 1
		 specialFieldRs.MoveNext
	  wend
	  
  
}
//��ʼ������(רҵ���༶)��Ϣ��¼����
 function initSCArray() {
   
	  i = 0
    sql = "select count(*) as count from [classInfo]"
	  set countRs = conn.Execute(sql)
    count = countRs("count") '�õ��ܵļ�¼��
	  
    sql = "select classSpecialFieldNumber,classNumber,className from [classInfo]"
	  set SpecialFieldClassRs = conn.Execute(sql)
	  
	  while not SpecialFieldClassRs.eof
	    Response.Write "specialField_class[" & i & "]='" & SpecialFieldClassRs("classSpecialFieldNumber") & ":" & SpecialFieldClassRs("classNumber") & ":" & SpecialFieldClassRs("className") & "';" & vbCrLf
		  i = i + 1
		  SpecialFieldClassRs.MoveNext
	  wend
  
}
//��ѡ��ͬ��רҵ��Ϣʱ��Ҫ���¸�רҵ�µ����а༶��Ϣ
function changeSpecialField() {
  var searchSpecialField; //Ҫ������רҵ
  var eachSpecialFiled; //ÿ����¼��רҵ
  var eachClassInfo; //ÿ���༶����Ϣ
  var eachClassNumber; 	//ÿ���༶�İ༶���
  var eachClassName; //��¼ÿ���༶������
  var indexOfSplit; // :�ŷָ���ŵ�λ��
  var innerHtmlText;
  var oOption; 
  var index;
  innerHtmlText = "";
  searchSpecialField = document.all.studentSpecialField.value;
  initSCArray(); //��ʼ��רҵ���༶��Ϣ����
  index = document.all.studentClass.length
  for(;index>0;index--) {
    document.all.studentClass.remove(index);
  }
  for(var i=0;i<specialField_class.length;i++) {
    indexOfSplit = specialField_class[i].indexOf(":"); //�õ�:�ŷָ���ŵ�λ��
	  eachSpecialField = specialField_class[i].substr(0,indexOfSplit); //ȡ�õ�ǰ��¼��רҵ���
	  if(searchSpecialField == eachSpecialField) { //���רҵһ���ͰѰ༶��Ϣȡ��������༶��������
	    eachClassInfo = specialField_class[i].substr(indexOfSplit+1);
		   indexOfSplit = eachClassInfo.indexOf(":");
			 eachClassNumber = eachClassInfo.substr(0,indexOfSplit); //ȡ�øü�¼�İ༶���
			 eachClassName = eachClassInfo.substr(indexOfSplit+1); //ȡ�øü�¼�İ༶����
		   oOption = document.createElement("OPTION");
		   document.all.studentClass.options.add(oOption);
			  oOption.innerText = eachClassName;
       oOption.value = eachClassNumber;
	  }
  }
}
//��ѡ��ͬ��ѧԺ��Ϣʱ��Ҫ���¸�ѧԺ�µ�רҵ��Ϣ
function changeCollege() {
  var searchCollege; //Ҫ������ѧԺ
  var eachCollege; //ÿ����¼��ѧԺ
  var eachSpecialFieldInfo; //ÿ��רҵ����Ϣ
  var eachSpecialFieldNumber; 	//ÿ����¼��רҵ���
  var eachSpecialFieldName; //��¼ÿ��רҵ������
  var indexOfSplit; // :�ŷָ���ŵ�λ��
  var innerHtmlText;
  var oOption; 
  var index;
  innerHtmlText = "";
  searchCollege = document.all.studentCollege.value;
  initCSArray(); //��ʼ��ѧԺ��רҵ��Ϣ����
  index = document.all.studentSpecialField.length
  for(;index>0;index--) {
    document.all.studentSpecialField.remove(index);
  }
  for(var i=0;i<college_specialField.length;i++) {
    indexOfSplit = college_specialField[i].indexOf(":"); //�õ�:�ŷָ���ŵ�λ��
	  eachCollege = college_specialField[i].substr(0,indexOfSplit); //ȡ�õ�ǰ��¼��ѧԺ���
	  if(searchCollege == eachCollege) { //���ѧԺһ���Ͱ�רҵȡ��������רҵ��������
	    eachSpecialInfo = college_specialField[i].substr(indexOfSplit+1);
		   indexOfSplit = eachSpecialInfo.indexOf(":");
			 eachSpecialFieldNumber = eachSpecialInfo.substr(0,indexOfSplit); //ȡ�øü�¼��רҵ���
			 eachSpecialFieldName = eachSpecialInfo.substr(indexOfSplit+1); //ȡ�øü�¼��רҵ����
		   oOption = document.createElement("OPTION");
		   document.all.studentSpecialField.options.add(oOption);
			  oOption.innerText = eachSpecialFieldName;
       oOption.value = eachSpecialFieldNumber;
	  }
  }
  changeSpecialField();
}
</script>
</HEAD>
<BODY onLoad="changeCollege();">
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>ѧ����Ϣ����--&gt;ѧ����Ϣ���
			 </td>
	   </tr><br>
		<tr>
		   <td>��ѡ������ѧԺ:</td>
			 <td>
			   <select id=studentCollege onChange="changeCollege();">
				  <%
				    '�õ����е�ѧԺ��Ϣ
					  dim sqlString
					  set rsCollege = Server.CreateObject("ADODB.RecordSet")
					  sqlString = "select * from [collegeInfo]"
					  rsCollege.Open sqlString,conn,1,1
						'����ÿ��ѧԺ����Ϣ����ӵ������б���
						while not rsCollege.EOF
						  Response.Write "<option value='" & rsCollege("collegeNumber") & "'>" & rsCollege("collegeName") & "</option>"
						  rsCollege.MoveNext
						wend
				  %>
				 </select>
			 </td>
		 </tr>
		  <td>��ѡ������רҵ:</td>
		  <td>
		    <select id=studentSpecialField onChange="changeSpecialField();">
			   <option value="">��ѡ��רҵ</option>
			  </select>
		  </td>
		 <tr>
		 <tr>
			 <td>��ѡ�����ڰ༶:</td>
			 <td>
			   <select name=studentClass id=studentClass>
				  <option value="">��ѡ��༶</option>
				</select>
			 </td>
		 </tr>
	   <tr>
	     <td style="height: 26px">
		     ѧ��:</td><td><input type=text name=studentNumber size=18></td>
			 </td>
		 </tr>
		 <tr>
		  <td>ѧ������:</td><td><input type=text name=studentName size=20></td>
		 </tr>
		 <tr>
		   <td>�Ա�:</td>
			 <td>
			   <select name=studentSex>
				   <option value='��'>��</option>
					 <option value='Ů'>Ů</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>ѧ������:</td>
			 <td>
			   <input type=text name=studentBirthday width=77px>
				 <input class="submit" name="Button" onClick="seltime('studentBirthday');" style="width:30px" type="button" value="ѡ��">
			 </td>
		 </tr>
		 <tr>
		   <td>������ò:</td>
			 <td>
			   <select name="studentState">
				   <option value='��Ա'>��Ա</option>
					 <option value='��Ա'>��Ա</option>
					 <option value='�ϰ���'>�ϰ���</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>��½����:</td>
			 <td><input type=text name=studentPassword size=20></td>
		 </tr>
		 
		 <tr>
			  <td>��Ƭ·��:</td>
			  <td><input type="text" name=photoAddress size=20 readonly>*���������ϴ���Ƭ,������Զ�����·��</td>
			</tr>
			<tr> 
       <td>��Ƭ�ϴ���</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  <tr>
		    <td>��ͥ��ַ:</td>
			  <td><input type=text name=studentAddress size=50></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=studentMemo></textarea></td>
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
