<!-- #include file="System/config.asp" -->
<HTML>
	<HEAD>
		<title><%=systemName%></title>
		<LINK href="css/style.css" type="text/css" rel="stylesheet">
		<script language="javascript">
			function check()
			{
				if(document.Form1.username.value == "")
				{
					alert("�������û��ʺ�");
					document.Form1.username.focus();
					return false;
				}
				else if(document.Form1.password.value == "")
				{
					alert("����������");
					document.Form1.password.focus();
					return false;
				}
				return true;
			}
		</script>
	</HEAD>
	<body>
		<form name="Form1" method="post"  action="checkLogin.asp" onsubmit="return check();">
			<TABLE id="Table1" width="80%" border="0" align=center>
				<tr>
					<td style="height: 14px" align="center">
                       </td>
				</tr>
				<br />
				<TR>
					<TD vAlign="middle" align="center">
						<TABLE class="tbl" id="Table2" cellSpacing="0" cellPadding="4" width="280" align="center"
							bgColor="#d6ebff" border="0">
							<TR>
								<TD class="bottom" align="center" bgColor="#52beef" colSpan="2" height="35">ϵͳ��½</TD>
							</TR>
							<TR>
								<TD class="bottom" align="center" colSpan="2"></TD>
							</TR>
							<TR>
								<TD class="br" style="HEIGHT: 33px" align="right" width="41%">�������û�����</TD>
								<TD class="bottom" style="HEIGHT: 33px" align="left" width="59%"><input type=text name=username size=20 value="admin"></TD>
							</TR>
							<TR>
								<TD class="br" align="right" width="41%">���������룺</TD>
								<TD class="bottom" align="left"><input type="password" name=password value="admin" size=20></TD>
							</TR>
							<TR>
								<TD class="br" align="right" width="41%">��ѡ����ݣ�</TD>
								<TD class="bottom" align="left">
								  <select name=identity>
								     <option value="student">ѧ��</option>
										 <option value="teacher">��ʦ</option>
										 <option value="admin" selected>����Ա</option>
								  </select>
							  </td>
							</TR>
							<TR>
								<TD align="center" colSpan="2" height="40">&nbsp;
									<input type="submit" value="��½">&nbsp;&nbsp;<input class="searchButton" id="btnExit" onclick="window.close();" type="button" value="�˳�"name="btnExit">
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<br />
				<tr>
					<td>
						<table align=center width=100% cellspacing=0 cellpadding=0>
	            <tr>
					    
	            </tr>
            </table>
					</td>
				</tr>
			</TABLE>
		</form>
	</body>
</HTML>