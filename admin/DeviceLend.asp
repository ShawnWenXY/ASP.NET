<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '���������Ϣ
  dim errMessage
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'���Ҫ�����豸�����
	if Request("submit") <> "" then
	  if Request("deviceName") = "" then
	    errMessage = "��û��ѡ���豸����!"
		end if
		if Request("studentNumber") = "" then
		  errMessage = "��û����дѧ��"
		end if
		sqlString = "select * from [studentInfo] where studentNumber='" & Request("studentNumber") & "'"
		set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
		studentInfoRs.Open sqlString,conn,1,1
		if studentInfoRs.EOF then
		  errMessage = "�Բ���,�������ѧ�Ų�����"
		end if
		studentInfoRs.Close
		if Request("deviceName") <> "" then
		  '��ѯ���豸������Ƿ����0
		  sqlString = "select * from [deviceStockInfo] where deviceName='" & Request("deviceName") & "'"
		  set deviceStockInfoRs = Server.CreateObject("ADODB.RecordSet")
		  deviceStockInfoRs.Open sqlString,conn,1,1
		  if CInt(deviceStockInfoRs("deviceStock")) <= 0 then
		    errMessage = "���豸û�п����"
		  end if
		  deviceStockInfoRs.Close
		end if
		'�����֤û�д���������µ��豸ʹ����Ϣ�ļ���
		if errMessage = "" then
		  sqlString = "select * from [deviceUseInfo]"
		  set deviceUseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  deviceUseInfoRs.Open sqlString,conn,1,3
		  deviceUseInfoRs.AddNew
		  deviceUseInfoRs("deviceName") = Request("deviceName")
		  deviceUseInfoRs("deviceTypeId") = Request("deviceTypeId")
		  deviceUseInfoRs("studentNumber") = Request("studentNumber")
		  deviceUseInfoRs("useBeginTime") = Now
		  deviceUseInfoRs("isReturn") = 0
		  deviceUseInfoRs.Update
		  deviceUseInfoRs.Close
		  '�����豸�Ŀ������1
		  sqlString = "update [deviceStockInfo] set deviceStock = deviceStock - 1 where deviceName='" & Request("deviceName") & "'"
		  conn.Execute(sqlString)
		  Response.Write "<script>alert('�豸ʹ�õǼ����!');</script>"
		else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
		end if
	end if
%>
<HEAD>
	<Title>�豸���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	var deviceType_deviceName = new Array();
	//��ʼ������(�豸���ͣ��豸������Ϣ)��Ϣ��¼����
  function initArray() {
  <%
    sql = "select * from deviceStockInfo" 
	  set deviceStockInfoRs = conn.Execute(sql)
	  i = 0
	  while not deviceStockInfoRs.eof
	    Response.Write "deviceType_deviceName[" & i & "]='" & deviceStockInfoRs("deviceTypeId") & ":" & deviceStockInfoRs("deviceName") & "';" & vbCrLf
		  i = i + 1
		  deviceStockInfoRs.MoveNext
	  wend
  %>
 }
 //��ѡ��ͬ���豸���ʱ��ʾ������µ��豸��Ϣ
function changeDeviceType() {
  var searchDeviceTypeId; //Ҫ�������豸����
  var eachDeviceTypeId; //ÿ����¼���豸����
  var eachDeviceName; //��¼ÿ���豸������
  var indexOfSplit; // :�ŷָ���ŵ�λ��
  var innerHtmlText;
  var oOption; 
  var index;
  innerHtmlText = "";
  searchDeviceTypeId = document.all.deviceTypeId.value;
  initArray(); //��ʼ����Ϣ����
  index = document.all.deviceName.length
  for(;index>0;index--) {
    document.all.deviceName.remove(index);
  }
  for(var i=0;i<deviceType_deviceName.length;i++) {
    indexOfSplit = deviceType_deviceName[i].indexOf(":"); //�õ�:�ŷָ���ŵ�λ��
	  eachDeviceTypeId = deviceType_deviceName[i].substr(0,indexOfSplit); //ȡ�õ�ǰ��¼���豸���ͱ��
	  if(searchDeviceTypeId == eachDeviceTypeId) { //����豸����һ���Ͱ��豸��Ϣȡ���������豸��Ϣ��������
	    eachDeviceName = deviceType_deviceName[i].substr(indexOfSplit+1);
		   oOption = document.createElement("OPTION");
		   document.all.deviceName.options.add(oOption);
			  oOption.innerText = eachDeviceName;
       oOption.value = eachDeviceName;
	  }
  }
}
	</script>
</HEAD>
<BODY>
<br>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=500 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�豸��Ϣ����--&gt;�豸����
			 </td>
	   </tr>
		 <tr>
		   <td width=100px align="right">�豸����:</td>
			 <td>
			   <select name=deviceTypeId onchange="changeDeviceType();">
				  <option value=''>��ѡ���豸����</option>
				   <%
					   sqlString = "select * from [deviceTypeInfo]"
						 set deviceInfoRs = Server.CreateObject("ADODB.RecordSet")
						 deviceInfoRs.Open sqlString,conn,1,1
						 while not deviceInfoRs.EOF
						   Response.Write "<option value='" & deviceInfoRs("deviceTypeId") & "'>" & deviceInfoRs("deviceTypeName") & "</option>"
							deviceInfoRs.MoveNext
						 wend
						 deviceInfoRs.Close
					 %>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100 align="right">�豸����:</td>
			 <td>
			   <select name=deviceName>
				   <option value="">��ѡ���豸</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">ѧ��:</td>
			 <td>
			   <input type=text name=studentNumber size=20>
			 </td>
		 </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input name="submit"  type="submit" value="���õǼ�"> &nbsp;
				  <input type="reset" value="ȡ��"></td>
      </tr>
	 </table>
 </form>
</body>
</html>