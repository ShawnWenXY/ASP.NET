<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/md5.asp"--> 
<%
  'errMessage保存错误信息
  dim errMessage
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"

	end if
	'如果管理员添加了新的学生信息并提交
	if Request("submit") <> "" then
	  '如果没有选择班级
	  if Request("studentClass") ="" then
	    errMessage = "请选择学生所在的班级!"
	  end if
	  '如果学号没有输入
	  if Request("studentNumber") = "" then
	    errMessage = "请输入学生的学号!"
	  end if
	  '如果学生的登陆密码没有输入
	  if Request("studentPassword") = "" then
	    errMessage = "请输入学生的登陆密码!"
		end if
	  if errMessage <> "" then
	    Response.Write "<script>alert('" & errMessage & "');</script>"
		else
	    '将学生个人信息加入的数据库中
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
	    '如果管理员选择了学生的生日
	    if Request("studentBirthday") <> "" then
	      studentInfoRs("studentBirthday") = CDate(Request("studentBirthday"))
		  else
		    studentInfoRs("studentBirthday") = CDate("1900-1-1")
		  end if
		  studentInfoRs("studentAddress") = Request("studentAddress")
		  studentInfoRs("studentMemo") = Request("studentMemo")
		  studentInfoRs.Update
		  studentInfoRs.Close
		  Response.Write "<script>alert('学生信息添加成功!')</script>"
	  end if
	end if
%>

<HTML>
<HEAD>
	<Title>新学生信息添加</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	
	var college_specialField = new Array();
	var specialField_class = new Array();
	//初始化所有(学院－专业)信息记录数组
  function initCSArray() {
   
    Server.ScriptTimeout = "10"
    set conn=server.CreateObject("Adodb.Connection")
    Path="driver={SQL Server};server=.;uid=sa;pwd=123456;database=SchoolManage" 
    conn.open path
    dim i  '循环变量
	  i = 0
    sql = "select count(*) as count from [specialFieldInfo]"
	'set countRs = conn.Execute(sql)
   ' count = countRs("count") '得到总的记录数
    sql = "select * from specialFieldInfo" 
	  set specialFieldRs = conn.Execute(sql)
	  
	  while not specialFieldRs.eof
	    Response.Write "college_specialField[" & i & "]='" & specialFieldRs("specialCollegeNumber") & ":" &specialFieldRs("specialFieldNumber") & ":" &specialFieldRs("specialFieldName") & "';" & vbCrLf
		  i = i + 1
		 specialFieldRs.MoveNext
	  wend
	  
  
}
//初始化所有(专业－班级)信息记录数组
 function initSCArray() {
   
	  i = 0
    sql = "select count(*) as count from [classInfo]"
	  set countRs = conn.Execute(sql)
    count = countRs("count") '得到总的记录数
	  
    sql = "select classSpecialFieldNumber,classNumber,className from [classInfo]"
	  set SpecialFieldClassRs = conn.Execute(sql)
	  
	  while not SpecialFieldClassRs.eof
	    Response.Write "specialField_class[" & i & "]='" & SpecialFieldClassRs("classSpecialFieldNumber") & ":" & SpecialFieldClassRs("classNumber") & ":" & SpecialFieldClassRs("className") & "';" & vbCrLf
		  i = i + 1
		  SpecialFieldClassRs.MoveNext
	  wend
  
}
//当选择不同的专业信息时需要更新该专业下的所有班级信息
function changeSpecialField() {
  var searchSpecialField; //要搜索的专业
  var eachSpecialFiled; //每个记录的专业
  var eachClassInfo; //每个班级的信息
  var eachClassNumber; 	//每个班级的班级编号
  var eachClassName; //记录每个班级的名称
  var indexOfSplit; // :号分割符号的位置
  var innerHtmlText;
  var oOption; 
  var index;
  innerHtmlText = "";
  searchSpecialField = document.all.studentSpecialField.value;
  initSCArray(); //初始化专业－班级信息数组
  index = document.all.studentClass.length
  for(;index>0;index--) {
    document.all.studentClass.remove(index);
  }
  for(var i=0;i<specialField_class.length;i++) {
    indexOfSplit = specialField_class[i].indexOf(":"); //得到:号分割符号的位置
	  eachSpecialField = specialField_class[i].substr(0,indexOfSplit); //取得当前记录的专业编号
	  if(searchSpecialField == eachSpecialField) { //如果专业一样就把班级信息取出来加入班级下拉框中
	    eachClassInfo = specialField_class[i].substr(indexOfSplit+1);
		   indexOfSplit = eachClassInfo.indexOf(":");
			 eachClassNumber = eachClassInfo.substr(0,indexOfSplit); //取得该记录的班级编号
			 eachClassName = eachClassInfo.substr(indexOfSplit+1); //取得该记录的班级名称
		   oOption = document.createElement("OPTION");
		   document.all.studentClass.options.add(oOption);
			  oOption.innerText = eachClassName;
       oOption.value = eachClassNumber;
	  }
  }
}
//当选择不同的学院信息时需要更新该学院下的专业信息
function changeCollege() {
  var searchCollege; //要搜索的学院
  var eachCollege; //每个记录的学院
  var eachSpecialFieldInfo; //每个专业的信息
  var eachSpecialFieldNumber; 	//每个记录的专业编号
  var eachSpecialFieldName; //记录每个专业的名称
  var indexOfSplit; // :号分割符号的位置
  var innerHtmlText;
  var oOption; 
  var index;
  innerHtmlText = "";
  searchCollege = document.all.studentCollege.value;
  initCSArray(); //初始化学院－专业信息数组
  index = document.all.studentSpecialField.length
  for(;index>0;index--) {
    document.all.studentSpecialField.remove(index);
  }
  for(var i=0;i<college_specialField.length;i++) {
    indexOfSplit = college_specialField[i].indexOf(":"); //得到:号分割符号的位置
	  eachCollege = college_specialField[i].substr(0,indexOfSplit); //取得当前记录的学院编号
	  if(searchCollege == eachCollege) { //如果学院一样就把专业取出来加入专业下拉框中
	    eachSpecialInfo = college_specialField[i].substr(indexOfSplit+1);
		   indexOfSplit = eachSpecialInfo.indexOf(":");
			 eachSpecialFieldNumber = eachSpecialInfo.substr(0,indexOfSplit); //取得该记录的专业编号
			 eachSpecialFieldName = eachSpecialInfo.substr(indexOfSplit+1); //取得该记录的专业名称
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
		      <img src="../images/ADD.gif" width=14px height=14px>学生信息管理--&gt;学生信息添加
			 </td>
	   </tr><br>
		<tr>
		   <td>请选择所在学院:</td>
			 <td>
			   <select id=studentCollege onChange="changeCollege();">
				  <%
				    '得到所有的学院信息
					  dim sqlString
					  set rsCollege = Server.CreateObject("ADODB.RecordSet")
					  sqlString = "select * from [collegeInfo]"
					  rsCollege.Open sqlString,conn,1,1
						'遍历每个学院的信息并添加到下拉列表中
						while not rsCollege.EOF
						  Response.Write "<option value='" & rsCollege("collegeNumber") & "'>" & rsCollege("collegeName") & "</option>"
						  rsCollege.MoveNext
						wend
				  %>
				 </select>
			 </td>
		 </tr>
		  <td>请选择所在专业:</td>
		  <td>
		    <select id=studentSpecialField onChange="changeSpecialField();">
			   <option value="">请选择专业</option>
			  </select>
		  </td>
		 <tr>
		 <tr>
			 <td>请选择所在班级:</td>
			 <td>
			   <select name=studentClass id=studentClass>
				  <option value="">请选择班级</option>
				</select>
			 </td>
		 </tr>
	   <tr>
	     <td style="height: 26px">
		     学号:</td><td><input type=text name=studentNumber size=18></td>
			 </td>
		 </tr>
		 <tr>
		  <td>学生姓名:</td><td><input type=text name=studentName size=20></td>
		 </tr>
		 <tr>
		   <td>性别:</td>
			 <td>
			   <select name=studentSex>
				   <option value='男'>男</option>
					 <option value='女'>女</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>学生生日:</td>
			 <td>
			   <input type=text name=studentBirthday width=77px>
				 <input class="submit" name="Button" onClick="seltime('studentBirthday');" style="width:30px" type="button" value="选择">
			 </td>
		 </tr>
		 <tr>
		   <td>政治面貌:</td>
			 <td>
			   <select name="studentState">
				   <option value='团员'>团员</option>
					 <option value='党员'>党员</option>
					 <option value='老百姓'>老百姓</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>登陆密码:</td>
			 <td><input type=text name=studentPassword size=20></td>
		 </tr>
		 
		 <tr>
			  <td>照片路径:</td>
			  <td><input type="text" name=photoAddress size=20 readonly>*请在下面上传照片,程序会自动生成路径</td>
			</tr>
			<tr> 
       <td>照片上传：</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  <tr>
		    <td>家庭地址:</td>
			  <td><input type=text name=studentAddress size=50></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=studentMemo></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认添加 ">
		      <input type="reset" value=" 重新填写 ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>

</HTML>
