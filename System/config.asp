<!-- #include virtual="DataBase/conn.asp" -->
<!--
/*ϵͳ������Ϣ��*/
CREATE TABLE [dbo].[config] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,														/*��Ϣ��¼���*/
	[systemName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,		/*ϵͳ����*/
	[author] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,				/*��������*/
	[email] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,					/*����Email*/
	[qq] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,						/*����qq*/
	[phone] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,					/*���ߵ绰*/
	[pageSize] [int] NULL ,																					/*ÿҳ��ʾ��¼����*/
	[canSelect] [int] NULL ,																				/*�Ƿ񿪷�ѡ��*/
	[termId] [int] NULL ,																						/*ѡ��������Ⱥ�ѧ��*/
	[selectStartTime] [datetime] NULL ,															/*ѡ�ο�ʼʱ��*/
	[selectEndTime] [datetime] NULL 																/*ѡ�ν���ʱ��*/
) ON [PRIMARY]
-->
<%
  '��ϵͳ���ݿ��ж�ȡϵͳ������
  dim systemName,author,email,qq,phone,pageSize,canSelect,termId,selectStartTime,selectEndTime
  set rsConfig=server.CreateObject("adodb.recordset")
	sqlString="select * from config where id = 1"
	rsConfig.open sqlString,conn,1,1
	'����ȡ��ϵͳ��Ϣ������ȫ�ֱ�����
	if not rsConfig.EOF then
	  systemName = rsConfig.Fields("systemName")
	  author = rsConfig.Fields("author")
	  email = rsConfig.Fields("email")
	  qq = rsConfig.Fields("qq")
	  phone = rsConfig.Fields("phone")
	  pageSize = CInt(rsConfig.Fields("pageSize"))
	  canSelect = CInt(rsConfig.Fields("canSelect"))
	  termId = rsConfig.Fields("termId")
	  selectStartTime = CDate(rsConfig.Fields("selectStartTime"))
	  selectEndTime = CDate(rsConfig.Fields("selectEndTime"))
	end if
	rsConfig.Close
%>
