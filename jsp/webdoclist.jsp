<%@ page contentType="text/html; charset=gb2312" %>
<%@ page language="java" import="java.sql.*" %>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href="cssjs/style.css">
<title>WebOffice ���߱༭WORD ,EXCEL���ĵ�������ʾ</title>
<script language=javascript>
<!--
function CreateNew(flag)
{
window.open('webdocedit.jsp?oper=new&flag='+flag,'','')
return 
	
}
-->
</script>
<body>
<center><br><h2><font color="#0099ff">
WebOffice ���߱༭WORD ,EXCEL���ĵ�������ʾ</font></h2><br><br>
    <input type=button value="�½�Word �ĵ�"  onclick="javascript:CreateNew(1)" style="width:150">
    <br>
    <input type=button value="�½�Excel�ĵ�"  onclick="javascript:CreateNew(2)" style="width:150">
    <br>
    <input type=button value="�½�PowerPoint�ĵ�"  onclick="javascript:CreateNew(3)" style="width:150"><br><br><br><b><font style="font-size:12pt;color:green">�Զ�ʶ��Office�ĵ�����</font><br><br>
    <table width=90% bgcolor=black cellpadding=1 cellspacing=1>
    <tr bgcolor=#cccccc><td align=center nowrap><b>�ļ����</b></td><td><b>�ļ���</b></td><td><b>����</b></td><td ><b>�ļ���С</b></td><td align=center><b>����</b></td></tr>
    <%
	try
	{
	 	String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver";		
		//Class.forName(DBDriver).newInstance();
		Class.forName(DBDriver);
		
	}
	catch(java.lang.ClassNotFoundException e)
	{
		System.err.println("Conn: " + e.getMessage()); 
		out.println( "err:" + e.getMessage());
	}

	try
	{  	 
		//String ConnStr = "jdbc:odbc:webofficedsn"; 
		String rootpath=application.getRealPath("/");
		String ConnStr="jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ="+rootpath+"\\database\\weboffice.mdb";			
		Connection conn = DriverManager.getConnection(ConnStr); 
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery("select * from dt_document order by o_pkid DESC");			
		while(rs.next())
		{
			String strid = rs.getString("o_pkid");
			String strName = rs.getString("o_name");
			String strflsid = rs.getString("o_flsid");
			String strSize = rs.getString("o_size");
			int flag =Integer.parseInt(rs.getString("o_flag"));
			String strPdf = rs.getString("o_pdf");
			String strHtml = rs.getString("o_html");
			if (strid == null)strid="";
			if (strName == null)strName="";
			if (strflsid == null)strflsid="";
			if (strSize == null)strSize="";
			if (strPdf == null)strPdf="";
			if (strHtml == null)strHtml="";
			String strFlag="";
			switch(flag)
			{
				case 1:
					strFlag = "<a href=\"WebDocEdit.jsp?oper=edit+id="+ strid +"\" target=_blank><img alt=\"WORD�ĵ� ��"+ strName +".doc\" src=\"images/doc.gif\" border=0></a>";
					break;
				case 2:
					strFlag = "<a href=\"WebDocEdit.jsp?oper=edit+id="+ strid +"\" target=_blank><img alt=\"EXCEL�ĵ� ��"+ strName +".xls\" src=\"images/xls.gif\" border=0></a>";
					break;
				case 3:
					strFlag = "<a href=\"WebDocEdit.jsp?oper=edit+id="+ strid +"\" target=_blank><img alt=\"PowerPoint�ĵ� ��"+ strName +".ppt\" src=\"images/ppt.gif\" border=0></a>";
					break;
				default:
				 	strFlag = "<a href=\"WebDocEdit.jsp?oper=edit+id="+ strid +"\" target=_blank><img alt=\"Ĭ��ΪWORD�ĵ� ��"+ strName +".doc\" src=\"images/doc.gif\" border=0></a>";
			}
			if (strPdf.equals("1"))
			{
				 strPdf = "<a href=\"pdf/"+ strflsid  +".pdf\" target=_blank><img  alt=\"PDF�ĵ� ��"+ strName +".pdf\" src=\"images/pdf.gif\" border=0></a>";
			}
			else
			{
				strPdf="";
			}
 			if (strHtml.equals("1"))
			{
				strHtml = "<a href=\"html/"+ strflsid +".html\" target=_blank><img  alt=\"HTML�ĵ� ��"+ strName +".html\"  src=\"images/htm.gif\" border=0></a>";
			}	else
			{
				strHtml="";
			}
			String outStr = "<tr><td bgcolor=white align=center>"+strid+"</td><td bgcolor=white>"+strName+ "</td><td bgcolor=white>"+ strFlag +"&nbsp;"+ strPdf +"&nbsp;"+ strHtml +"</td><td bgcolor=white>"+strSize+"</td><td bgcolor=white align=center><a href=\"webdocread.jsp?oper=read&id="+ strid +"\" target=_blank>�Ķ�</a>&nbsp;&nbsp;<a href=\"webdocedit.jsp?oper=edit&id="+ strid +"\" target=_blank>�༭</a></td></tr>";
			out.println(outStr);
		}		
		rs.close();		
	}
	catch(SQLException e)
	{ 
		out.println(e.getMessage());
		System.err.println("QueryErr: " + e.getMessage()); 
	}	
    %>
    </table>
    </body>
    </html>
