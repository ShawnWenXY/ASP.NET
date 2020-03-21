<!--#include file="../Database/conn.asp"-->
<%
	'����ѧԺ��ŵõ�ѧԺ����
	Function GetCollegeNameByNumber(colleageNumber)
	  dim sqlString,collegeName
	  sqlString = "select collegeName from [collegeInfo] where collegeNumber='" & colleageNumber & "'"
	  set collegeInfoRs = Server.CreateObject("ADODB.RecordSet")
	  collegeInfoRs.Open sqlString,conn,1,1
	  if not collegeInfoRs.EOF then
	    collegeName = collegeInfoRs("collegeName")
		else
		  collegeName = ""
		end if
		collegeInfoRs.Close
		GetCollegeNameByNumber = collegeName
	End Function
	
	'����רҵ��ŵõ�רҵ����
	Function GetSpecialFieldNameByNumber(specialFieldNumber)
	  dim sqlString,specialFieldName
	  sqlString = "select specialFieldName from [specialFieldInfo] where specialFieldNumber='" & specialFieldNumber & "'"
	  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
	  specialFieldInfoRs.Open sqlString,conn,1,1
	  if not specialFieldInfoRs.EOF then
	    specialFieldName = specialFieldInfoRs("specialFieldName")
		else
		  specialFieldName = ""
		end if
		GetSpecialFieldNameByNumber = specialFieldName
	End Function
	
	'���ݰ༶��ŵõ��༶����
	Function GetClassNameByNumber(classNumber)
	  dim sqlString,className
		sqlString = "select className from [classInfo] where classNumber='" & classNumber & "'"
		set classInfoRs = Server.CreateObject("ADODB.RecordSet")
		classInfoRs.Open sqlString,conn,1,1
		if not classInfoRs.EOF then
		  className = classInfoRs("className")
		else
		  className = ""
		end if
		GetClassNameByNumber = className
	End Function
	
	'����ѧ�ڱ�ŵõ�ѧ������
	Function GetTermnameById(termId)
	  dim sqlString,termName
	  sqlString = "select * from [termInfo] where termId=" & termId
	  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
	  termInfoRs.Open sqlString,conn,1,1
	  if not termInfoRs.EOF then
	    termName = termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "��" & termInfoRs("termUpOrDown")
	  else
	    termName = ""
		end if
		GetTermnameById = termName
	End Function
	
	'���ݽ�ʦ��ŵõ���ʦ������
	Function GetTeacherNameByNumber(teacherNumber)
	  dim sqlString,teacherName
	  sqlString = "select teacherName from [teacherInfo] where teacherNumber='" & teacherNumber & "'"
	  set teacherInfoRs = Server.CreateObject("ADODB.RecordSet")
	  teacherInfoRs.Open sqlString,conn,1,1
	  if not teacherInfoRs.EOF then
		  teacherName = teacherInfoRs("teacherName")
	  else
	    teacherName = ""
	  end if
	  GetTeacherNameByNumber = teacherName
	End Function
	
	'���ݰ༶�γ̱�ŵõ��༶�γ�����
	Function GetClassCourseNameByNumber(courseNumber)
	  dim sqlString,classCourseName
	  sqlString = "select * from [classCourseInfo] where courseNumber='" & courseNumber & "'"
	  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	  classCourseInfoRs.Open sqlString,conn,1,1
	  if not classCourseInfoRs.EOF then
	    classCourseName = classCourseInfoRs("courseName")
	  else
	    classCourseName = ""
		end if
		GetClassCourseNameByNumber = classCourseName
	End Function
	
	'����ѡ�޿γ̱�ŵõ�ѡ�޿γ̵�����
	Function GetPublicCourseNameByNumber(courseNumber)
	  dim sqlString,publicCourseName
	  sqlString = "select courseName from [publicCourseInfo] where courseNumber='" & courseNumber & "'"
	  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	  publicCourseInfoRs.Open sqlString,conn,1,1
	  if not publicCourseInfoRs.EOF then
		  publicCourseName = publicCourseInfoRs("courseName")
		else
		  publicCourseName = ""
		end if
		GetPublicCourseNameByNumber = publicCourseName
	End Function
	
	'����ѧ�ŵõ�ѧ��������
	Function GetStudentNameByNumber(studentNumber)
	  dim sqlString,studentName
	  sqlString = "select studentName from [studentInfo] where studentNumber='" & studentNumber & "'"
	  set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
	  studentInfoRs.Open sqlString,conn,1,1
	  if not studentInfoRs.EOF then
	    studentName = studentInfoRs("studentName")
		else
		  studentName = ""
	  end if
	  GetStudentNameBYNumber = studentName
	End Function
	
	'����ѡ�޿γ̱�ŵõ��ÿγ̵��Ͽν�ʦ������
	Function GetTeacherNameByPublicCourseNumber(courseNumber)
	  dim sqlString,teacherName
	  sqlString = "select teacherName from [teacherInfo],[publicCourseInfo],[publicCourseTeach] where [publicCourseInfo].courseNumber='" & courseNumber & "' and [publicCourseInfo].courseNumber = [publicCourseTeach].courseNumber and [publicCourseTeach].teacherNumber = [teacherInfo].teacherNumber"
	  set rs = Server.CreateObject("ADODB.RecordSet")
	  rs.Open sqlString,conn,1,1
	  if not rs.EOF then
	    teacherName = rs("teacherName")
		else
		  teacherName = ""
		end if
		GetTeacherNameByPublicCourseNumber = teacherName
	End Function
	
	'����ѧ�ŵõ���ѧ�����ڵ�רҵ
	Function GetSpecialFieldNumberByStudentNumber(studentNumber)
	  dim sqlString,specialFieldNumber
	  sqlString = "select specialFieldNumber from [specialFieldInfo],[classInfo],[studentInfo] where [studentInfo].studentNumber='" & studentNumber & "' and [studentInfo].studentClassNumber = [classInfo].classNumber and [classInfo].classSpecialFieldNumber = [specialFieldInfo].specialFieldNumber"
	  set rs = Server.CreateObject("ADODB.RecordSet")
	  rs.Open sqlString,conn,1,1
	  if not rs.EOF then
	    specialFieldNumber = rs("specialFieldNumber")
		else
		  specialFieldNumber = ""
		end if
		GetSpecialFieldNumberByStudentNumber = specialFieldNumber
	End Function
  
  '����ѧ�ŵõ��༶���
  Function GetClassNumberByStudentNumber(studentNumber)
    dim sqlString,classNumber
	   sqlString = "select studentClassNumber from [studentInfo] where studentNumber='" & studentNumber & "'"
		 set rs = Server.CreateObject("ADODB.RecordSet")
		 rs.Open sqlString,conn,1,1
		 if not rs.EOF then
		   classNumber = rs("studentClassNumber")
		 else
		   classNumber = ""
		 end if
		 GetClassNumberByStudentNumber = classNumber
  End Function
  
  '��������:�����豸���͵õ��豸����
	Function GetDeviceTypeNameById(deviceTypeId)
	  dim sqlString,deviceTypeName
	  sqlString = "select deviceTypeName from [deviceTypeInfo] where deviceTypeId=" & deviceTypeId
	  set deviceTypeInfoRs = Server.CreateObject("ADODB.RecordSet")
	  deviceTypeInfoRs.Open sqlString,conn,1,1
	  if not deviceTypeInfoRs.EOF then
	    deviceTypeName = deviceTypeInfoRs("deviceTypeName")
		else
		  deviceTypeName = ""
		end if
		GetDeviceTypeNameById = deviceTypeName
	End Function
%>