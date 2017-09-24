<%@ page contentType="text/html; charset=gb2312" %>
<%@ page import="java.io.*,java.text.*,java.util.*,java.sql.*,javax.servlet.*,javax.servlet.http.*" %>
<html>
<head>
<%
request.setCharacterEncoding("GB2312");
String strHeight = "80%";
String strTitlebar = "1";
String strToolbar = "1";
String strHost = request.getServerName();
String strURL = request.getRequestURI();
strURL = strURL.substring(0,strURL.lastIndexOf("/")+1);
String strDefaultRoot = "http://" + strHost + strURL;
String strOpenUrl = "http://" + strHost  + strURL + "writefile.jsp";
String strPdfSaveUrl ="http://" + strHost + strURL + "pdf.jsp";
String strHTMLSaveUrl ="http://" + strHost + strURL + "html.jsp";
String strppFileSaveUrl ="http://" + strHost + strURL + "ppfile.jsp";
strURL ="http://" + strHost + strURL + "uploadedit.jsp"; 
String office =  request.getParameter("office");
String strflsid =  request.getParameter("flsid");
String strnum =  request.getParameter("num");
String strfname =  request.getParameter("fname");
String strfcreator =  request.getParameter("fcreator");
String tselect =  request.getParameter("tselect");
String strflag =  request.getParameter("flag");
String stroper =  request.getParameter("oper");
String strJsScript="";

String loadfun="";
if (office=="wps"){
	loadfun="WebOpen('1');";
}
else{
	loadfun="WebOpen('');";
}

String strid = request.getParameter("id");
String strSQL = null;

String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 		   
String rootpath=application.getRealPath("/");
String ConnStr="jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ="+rootpath+"\\database\\weboffice.mdb";			
//String ConnStr = "jdbc:odbc:webofficedsn";
Connection conn = null; 
ResultSet rs = null; 
Statement stmt = null;
if (stroper.compareTo("read")==0)
{
	  strSQL = "select * from dt_document where  o_pkid="+strid;
	  try 
		{ 
			
			Class.forName(DBDriver).newInstance(); 
		} 
		catch(java.lang.ClassNotFoundException e)
		{ 
			System.err.println("Conn(): " + e.getMessage()); 
		}

	try
		{ 
		conn = DriverManager.getConnection(ConnStr); 
		stmt = conn.createStatement();
		rs = stmt.executeQuery(strSQL);
			if(rs.next())
			{
			 
			  try
				  {
					strfname = rs.getString("o_name");
					strfcreator = rs.getString("o_creator");
					strnum = rs.getString("o_number");
					strflsid = rs.getString("o_flsid");
					strflag = rs.getString("o_flag");
					if (strfname == null)strfname="";
					if (strfcreator == null)strfcreator="";
					if (strnum == null)strnum="";
					if (strflsid == null)strflsid="";
					if (strflag == null)strflag="";
				  }
				  catch(Throwable e)
				  {
					System.out.println(e.toString());
					throw new ServletException(e.toString());
				  }


			}
			stmt.close();
			rs.close();
			conn.close();
		} 
	catch(SQLException e)
		{ 
		System.err.println("QueryErr: " + e.getMessage()); 
		 
		}
 
	strOpenUrl = strOpenUrl+ "?id="+strid +"&operfile=office&file="+strflsid;
	strURL = strURL + "?oper=edit&flsid="+ strflsid + "&flag="+strflag;
	strJsScript = strJsScript + "<"+"script language=javascript>flag='1"+ strflag +"';<"+"/script>";
	strTitlebar = "0";
	strToolbar = "0";
}
 



SimpleDateFormat formatter=new java.text.SimpleDateFormat("yyyyMMddHHmmss"); 
java.util.Date currentTime=new java.util.Date(); 
String crrtime=formatter.format(currentTime);

if (stroper.compareTo("new")==0)
{ 
		strflsid = crrtime;
		/*
		switch(Integer.parseInt(strflag))	 
		{ 
			case 2:
			strflsid = strflsid+".xls";
			break;
			case 3:
			strflsid = strflsid+".ppt";
			break;
			default:
			strflsid = strflsid+".doc";
				
		}
		*/		 
		strURL = strURL+"?oper=new&flsid="+strflsid ; 
		strnum = crrtime;		
		stroper = "addnew";
		if( tselect!=null && tselect.compareTo("0")==0 && tselect!=""){
			strJsScript = strJsScript + "<script language=javascript>flag='0';</script>";
			//	out.println( "<script language=javascript>flag='0';</script>");
				strOpenUrl = strOpenUrl + "?oper=new&id=" + tselect;
		}else
		{
			if(strflag!=""){
		 
				//out.println(  "<script language=javascript>flag='"+strflag+"';</script>");
				strJsScript = strJsScript + "<script language=javascript>flag='"+strflag+"';</script>";

				
				}
		}
		
if (strfname==null || strfname==""){ strfname="测试文档"+ crrtime;}
 if(strnum=="" || strnum ==null)strnum=crrtime;
			 
}
 

%>


<title>WebOffice办公文档控件演示版：<%=strfname%></title>
<link rel=stylesheet type=text/css href="cssjs/style.css">
<script language=javascript>
var strRoot;
var strOpenUrl;
var strURL;
var autoSave=1;
var pfile='<%=strflsid%>'; 
var flag='<%=strflag%>';
strOpenUrl = '<%=strOpenUrl%>';
strURL='<%=strURL%>';
strRoot = '<%=strDefaultRoot%>';
var strPdfSaveUrl='<%=strPdfSaveUrl%>';
var strHTMLSaveUrl='<%=strHTMLSaveUrl%>';
var strppFileSaveUrl='<%=strppFileSaveUrl%>';
</script><%=strJsScript%>
<script language=javascript src="cssjs/weboffice.js"></script>
</head>
<body topmargin=0 leftmargin=0 onload="javascript:<%=loadfun%>"><form action="WebDocEdit.jsp?oper=<%=stroper%>"  name=frm method="post" onsubmit="return WebSave()">
<table width=100% bgcolor=#cccccc cellpadding=1 cellspacing=1><input type=hidden value=<%=strflag%> name=flag><tr>
<td nowrap bgcolor=white valign=top style="padding-top:10px;"><div align=left>&nbsp;无法看到控件?&nbsp;<a href="http://www.officectrl.com/weboffice/weboffice.rar"><img src="images/weboffice-install.jpg" border=0></a>完成后重新打开本页面。</div><script language=javascript src="cssjs/webofficeocx.js" charset="utf-8"></script>

</td></tr></table>&nbsp;&nbsp;<font color=red>&copy;</font>2001-2016 All Rights Reserved!&nbsp;&nbsp;本站网址：<a href="http://www.officectrl.com/">http://www.officectrl.com</a><br></div><br><br></form>
</body></html>

