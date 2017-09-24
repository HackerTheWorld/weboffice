<%@ page contentType="text/html; charset=gb2312" %>
<%@ page import="java.io.*,java.text.*,java.util.*,java.sql.*,javax.servlet.*,javax.servlet.http.*" %>
<html>
<head>
<%
request.setCharacterEncoding("GB2312");
String strHeight = "100%";
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
	out.println("<"+"script language=javascript>flag='1"+ strflag +"';<"+"/script>");
	strOpenUrl = strOpenUrl+ "?id="+strid +"&operfile=office&file="+strflsid;
	strURL = strURL + "?oper=edit&flsid="+ strflsid + "&flag="+strflag;
	strHeight="500";
	strTitlebar = "0";
	strToolbar = "0";
}

if (stroper.compareTo("doedit")==0) 
{
		 
			 
	 strSQL = "update dt_document set o_number='"+strnum+"',o_name='"+strfname+"', o_flag="+strflag+" where o_flsid='"+strflsid+"'";
	
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
		stmt.executeUpdate(strSQL);
		stmt.close();
			conn.close();

		out.println("<script language=javascript>window.opener.location.reload();window.close();</script>"); 
	} 
	catch(SQLException e)
	{ 
		System.err.println("QueryErr: " + e.getMessage()); 
		 
	} 
	
	 return; 


	
}



if (stroper.compareTo("addnew")==0)
{
	strfcreator ="admin";
	 
	strSQL = "update dt_document set o_number='"+strnum+"',o_name='"+strfname+"',o_creator='"+strfcreator+"', o_flag="+strflag+" where o_flsid='"+strflsid+"'";
		
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
		stmt.executeUpdate(strSQL);
			stmt.close();
			
			conn.close();

		out.println( "<script language=javascript>window.opener.location.reload();window.close();</script>");
	} 
	catch(SQLException e)
	{ 
		System.err.println("QueryErr: " + e.getMessage()); 
		
	} 
 	return;
}

if (stroper.compareTo("edit")==0)
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
		strJsScript = strJsScript + "<"+"script language=javascript>flag='1"+ strflag +"';<"+"/script>";
	//out.println("<"+"script language=javascript>flag='1"+ strflag +"';<"+"/script>");
	strOpenUrl = strOpenUrl+ "?id="+strid +"&operfile=office&file="+strflsid;
	strPdfSaveUrl = strPdfSaveUrl+ "?flsid="+ strflsid;
	strHTMLSaveUrl = strHTMLSaveUrl+ "?flsid="+ strflsid;
	strppFileSaveUrl = strppFileSaveUrl+ "?flsid="+ strflsid;
	strURL = strURL + "?oper=edit&flsid="+ strflsid + "&flag="+strflag+"&id="+strid;
	stroper = "doedit";   
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
		
if (strfname==null || strfname==""){ strfname="�����ĵ�"+ crrtime;}
 if(strnum=="" || strnum ==null)strnum=crrtime;
			 
}
 

%>


<title>WebOffice�칫�ĵ��ؼ���ʾ�棺<%=strfname%></title>
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
<table width=100% bgcolor=#cccccc cellpadding=1 cellspacing=1><input type=hidden value=<%=strflag%> name=flag><tr><td width=80 valign=top bgcolor=white align=center><input type=hidden name=oper value=<%=stroper%> id=oper>
[<b style="color:green">�������</b>]<br>
<input type=button  class="button" value="ͬ������" style="width:80" onclick="WebHttpSave();"><br>
<input type=submit class="button"   value="�첽����" style="width:80"><br>
<input type=button class="button"   value="�ص��ĵ�"  onclick="WebDocReload()"><br>[<b style="color:green">����</b>]<br><input type=button  class="button" value="��������" onclick="frm.WebOffice.IsNotCopy=1;"><br>
<input type=button class="button" value="������" onclick="frm.WebOffice.IsNotCopy=0;"><br>
[<b style="color:green">�Զ�����</b>]<br>
<input type=button class="button" id=savbtn onclick="alert('��ǰ����ʹ�ؼ�ÿ��30���Զ������ĵ������Ե�,�ؼ����ڴ���...');MyTimer();this.value='�����ѿ���';clsbtn.style.display='';"  value="�����Զ�����" style="width:80"><input type=button id=clsbtn style="display:none;" class="button" onclick="autoSave=0;this.value='�Զ������ѹ�';savbtn.value='�����Զ�����';info.innerHTML='';this.style.display='none'" value="�ر��Զ�����" style="width:80"><font id=info style="display:none"></font>[<b style="color:green">�򿪱���</b>]<br>
<input type=button class="button" value="�򿪷�ʽһ" onclick="var a=document.all.WebOffice.WebLoadFile('http://www.officectrl.com/weboffice/temp/file1.doc',null);"><br>
<input type=button class="button" value="�򿪷�ʽ��" onclick="var a=document.all.WebOffice.WebLoadFile('http://www.officectrl.com/weboffice/temp/file1.doc','Word.Document');"><br>
<input type=button class="button" value="�򿪷�ʽ��" onclick="document.all.WebOffice.WebOpen('http://www.officectrl.com/weboffice/temp/file1.doc');"><br>
<input type=button class="button" value="�򿪷�ʽ��" onclick="document.all.WebOffice.WebUrl('http://www.officectrl.com/weboffice/temp/file1.doc','Word.Document');"><br>
<input type=button class="button" value="�½��ļ�EXCEL" onclick="var a=document.all.WebOffice.WebLoadFile('','xls');frm.flag.value=2;"><br>
<input type=button class="button" value="�½��ļ�WORD" onclick="document.all.WebOffice.WebLoadFile('','doc');frm.flag.value=1;"><br>
<input type=button class="button" value="�½��ļ�PPT" onclick="var a=document.all.WebOffice.WebLoadFile('','ppt');frm.flag.value=3;"><br>
<input type=button class="button"   value="���浽C:\a.doc"  onclick="WebSaveC()">
<input type=button class="button"   value="�򿪱����ļ�"  onclick="WebOpenLocal()"><br>
<input type=button class="button"   value="��Ϊ�����ļ�"  onclick="WebSaveLocal()"><br>
<input type=button class="button"   value="ָ��ҳ����"  onclick="document.all.WebOffice.SetPageAs('c:\\b.doc',1,0);alert('�ѽ��ĵ���1ҳ���������Ϊc�̸�Ŀ¼�µ�b.doc�ļ�');"><br>
<!--<input type=button class="button"   value="����ָ���ļ�"  onclick="document.all.WebOffice.DownloadFile('http://www.officectrl.com/weboffice/temp/file1.doc','c:\\a.doc');alert('�ѽ�http://www.officectrl.com/weboffice/temp/file1.doc\n���ص�c�̸�Ŀ¼�µ�a.doc');"><br>
--><input type=button class="button"   value="ɾ���ļ�"  onclick="document.all.WebOffice.DeleteLocalFile('c:\\a.doc');alert('�ѽ�c�̸�Ŀa.docɾ��');"><br>

[<b style="color:green">�ĵ�ת��</b>]<br>
<%if(strflag.equals("1")){%>
<input type=button class="button"   value="����Զ��PDF"  onclick="WebSaveRemotePdf();"><br>
<input type=button class="button"   value="����Զ��HTML"  onclick="WebSaveRemoteHTML()" id="idWordTable"><br>
<input type=button class="button"   value="���汾��PDF"  onclick="WebSaveLocalPdf()"><br>
<input type=button class="button"   value="���汾��HTML"  onclick="WebSaveLocalHTML()"><br>
<%}%><%if(strflag.equals("2")){%><input type=button class="button"   value="���汾��PDF"   onclick="WebSaveXLSLocalPDF();"><br>
<input type=button class="button"   value="���汾��HTML"   onclick="WebSaveXLSLocalHTML();"><br>
<input type=button class="button"   value="����Զ��PDF"   onclick="WebSaveXLSAsPDF();"><br>
<input type=button class="button"   value="����Զ��HTML"   onclick="WebSaveXLSAsHTML();"><br>
<input type=button class="button"   value="����Զ��MHT"   onclick="WebSaveXLSAsMHT();"> <br>
<%}%><%if(strflag.equals("3")){%><input type=button class="button"   value="���汾��PDF"   onclick="WebSavePPTLocalPDF();"><br>
<input type=button class="button"   value="���汾��ͼƬ"   onclick="WebSavePPTLocalJPG();"><br>
<input type=button class="button"   value="����Զ��PDF"   onclick="WebSavePPTAsPDF();"><br>
<input type=button class="button"   value="����Զ��HTML"   onclick="WebSavePPTAsHTML();"><br>
<%}%><input type=hidden  value="" id=field1 name=field1>
<input type=hidden  value="" id=field2 name=field2>
<input type=hidden  value="" id=field3 name=field3>
<input type=hidden  value="" id=field4 name=field4>
<input type=hidden  value="<%="http://" +  request.getServerName() + request.getRequestURI() + "?" + request.getQueryString()%>" id=field5 name=field5>
<input type=hidden  value="<%=strDefaultRoot%>" id=field6 name=field6>
[<b style="color:green">��ӡ����</b>]
<input type=button class="button"   value="��ӡԤ��"  onclick="WebDocPrintPreView()">
<input type=button class="button"   value="���ú��ӡһ"  onclick="WebDocPrint()">
<input type=button class="button"   value="���ú��ӡ��"  onclick="WebDocPrint2()">
<input type=button class="button"   value="ֱ�Ӵ�ӡ"  onclick="WebPrintDirc()">
[<b style="color:green">��������</b>]<br>
<input type=button class="button"   value="�رձ�����"   onclick="WebTitlebar(false)"><br>
<input type=button class="button"   value="�򿪱�����"   onclick="WebTitlebar(true)"><br>
<input type=button class="button"   value="�رղ˵���"   onclick="WebMenubar(false)"><br>
<input type=button class="button"   value="�򿪲˵���"   onclick="WebMenubar(true)"><br>
<input type=button class="button"   value="�رչ�����"   onclick="WebToolbar(false)"><br>
<input type=button class="button" value="�򿪹�����"   onclick="WebToolbar(true)"><br>
<!--<input type=button class="button" value="��ֹ�ļ��˵�(�����½����򿪡������)" onclick="alert('�˹��ܿ��Խ��ؼ��ļ��˵���Ĳ˵�����Ч��');document.all.WebOffice.SetMenuDisplay(0);"><br>
<input type=button class="button" value="����ϵͳʱ��" onclick="alert('��ȷ���󣬽�����ϵͳʱ������Ϊ��2016-09-02 12:10:10');document.all.WebOffice.SetCurrTime('2016-09-02 12:10:10');"><br>

<input type=button class="button" value="Office�汾" onclick="var a=document.all.WebOffice.GetOfficeVersion('Word.Document');alert(a);"><br>-->
[<b style="color:green">�ĵ���ͼ</b>]<br>
<input type=button class="button"   value="��ͨ��ͼ"  onclick="document.all.WebOffice.ShowView(1);"><br>
<input type=button class="button"   value="�����ͼ"  onclick="document.all.WebOffice.ShowView(2);"><br>
<input type=button class="button"   value="ҳ����ͼ"  onclick="document.all.WebOffice.ShowView(3);"><br>
<input type=button class="button"   value="��ӡԤ��"  onclick="document.all.WebOffice.ShowView(4);"><br>
<input type=button class="button"   value="������ͼ"  onclick="document.all.WebOffice.ShowView(5);"><br>
<input type=button class="button"   value="WEB��ͼ"  onclick="document.all.WebOffice.ShowView(6);"><br>
<input type=button class="button"   value="�Ķ���ͼ"  onclick="document.all.WebOffice.ShowView(7);"><br>
[<b style="color:green">�ĵ�����</b>]<br>
<input type=button class="button" value="Applicationʾ��" onclick="var app=document.all.WebOffice.GetApplication;app.Dialogs(163);"><br>
<input type=button class="button" value="ActiveDocumentʾ��" onclick="form.WebOffice.ActiveDocument.Application.Dialogs(129);"><br>
<input type=button class="button"   value="ȡ�û���·��"  onclick="var a=document.all.WebOffice.TempFilePath;alert(a);"><br>
</td><td  nowrap bgcolor=white  valign=top width=80 align=center>[<b style="color:green">ȫ����ʾ</b>]<br>
<input type=button class="button" value="����ȫ��" class="button"   onclick="document.all.WebOffice.MenuBars=1;document.all.WebOffice.FullScreenType=1;document.all.WebOffice.WebFullScreen();"><br>
<!--<input type=button class="button" value="IE��ȫ��" class="button"   onclick="document.all.WebOffice.MenuBars=1;document.all.WebOffice.FullScreenType=2;document.all.WebOffice.WebFullScreen();"><br>[<b style="color:green">��������</b>]<br>
--><input type=button class="button"   value="��дǩ��"  onclick="WebDocSignature()"><br>
<input type=button class="button"   value="��Ӳ�����"  onclick="WebSignature('1')"><br>
<input type=button class="button"   value="��Ӻ�ͬ��"  onclick="WebSignature('2')"><br>
<input type=button class="button"   value="���������"  onclick="WebSignature('3')"><br> 
[<b style="color:green">��ǩ/�滻</b>]<br>
<input type=button class="button" value="�滻�ı�" onclick="document.all.WebOffice.ReplaceText('a','b',0);"><br>
<input type=button class="button" value="�����ǩ" onclick="document.all.WebOffice.SetFieldValue('mark_1','test','::ADDMARK::');"><br>
<input type=button class="button" value="��ǩ��ɫ" onclick="document.all.WebOffice.SetFieldValue('mark_1','255','::SETCOLOR::');"><br>
<!--<input type=button class="button" value="�����ǩɫ" onclick="var vcolor= document.all.WebOffice.SetFieldValue('mark_1','','::GETCOLOR::');alert(vcolor);"><br>-->
<input type=button class="button" value="ɾ����ǩ" onclick="document.all.WebOffice.SetFieldValue('mark_1','','::DELMARK::');"><br>
<input type=button class="button" value="�����ǩ" onclick="document.all.WebOffice.SetFieldValue('mark_1','','::GETMARK::');"><br>
<input type=button class="button" value="���������ǩ" onclick="WebGetAllMark()"><br>
<input type=button class="button" value="�����׺�SetFieldValue" onclick="document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/temp/file1.doc','::FILE::');"><br>
<input type=button class="button" value="����ͼƬ" onclick="WebAddPic()"><br>
<input type=button class="button" value="���븡��ͼƬ" onclick="WebAddFloatPic()"><br>
<input type=button class="button" value="EXCEL����ͼƬSetFieldValue" onclick="document.all.WebOffice.SetFieldValue('mark_1','test','::ADDMARK::');document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/images/weboffice.jpg','::JPG::');"><br>
<input type=button class="button"   name="insertword"  value="����WORD�ĵ�"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',0);"><br>

[<b style="color:green">�ĵ�����</b>]<br>
<input type=button class="button" value="�ĵ���ȫ����" onclick="document.all.WebOffice.ProtectDoc(1,2,'123');"><br>
<input type=button class="button" value="����������޶�" onclick="document.all.WebOffice.ProtectDoc(1,0,'123');"><br>
<input type=button class="button" value="�����������ע" onclick="document.all.WebOffice.ProtectDoc(1,1,'123');"><br>
<input type=button class="button" value="�ĵ��������" onclick="document.all.WebOffice.ProtectDoc(0,0,'123');"><br>
<input type=button class="button"   value="Excel�������ݵ�Ԫ��" onclick="WebSheetsLock()"><br>
 [<b style="color:green">ҳ�����</b>]<br>
<input type=button class="button"   value="�ĵ�ҳ��"  onclick="WebDocumentPageCount()"><br>
<input type=button class="button"   value="����ҳü"  onclick="frm.WebOffice.ActiveDocument.ActiveWindow.ActivePane.View.SeekView=9;"><br>
<input type=button class="button"   value="����ҳ��"  onclick="frm.WebOffice.WebDialogs(294);"><br>
<input type=button class="button"   value="�����ַ�"  onclick="frm.WebOffice.WebDialogs(162);"><br>
<input type=button class="button"   value="����Ŀ¼"  onclick="frm.WebOffice.WebDialogs(171);"><br>
<input type=button class="button"   value="������"  onclick="frm.WebOffice.WebDialogs(129);"><br>
[<b style="color:green">Word�Ի���</b>]<br>
<input type=button class="button"   value="��Ŀ��������ToolsBulletsNumbers"  onclick="frm.WebOffice.WebDialogs(196);"><br>
<input type=button class="button"   value="��������TableSort"  onclick="frm.WebOffice.WebDialogs(199);"><br>
<input type=button class="button"   value="����ͳ��ToolsWordCount"  onclick="frm.WebOffice.WebDialogs(228);"><br>
<input type=button class="button"   value="ȡ��ѡ��ShrinkSelection"  onclick="frm.WebOffice.WebDialogs(236);"><br>
<input type=button class="button"   value="ѡ��ȫ��EditSelectAll"  onclick="frm.WebOffice.WebDialogs(237);"><br>
<input type=button class="button"   value="����ҳ��InsertPageField"  onclick="frm.WebOffice.WebDialogs(239);"><br>
<input type=button class="button"   value="��������InsertDateField"  onclick="frm.WebOffice.WebDialogs(240);"><br>
<input type=button class="button"   value="����ʱ��InsertTimeField"  onclick="frm.WebOffice.WebDialogs(241);"><br>
<input type=button class="button"   value="ҳ������FilePageSetup"  onclick="frm.WebOffice.WebDialogs(178);"><br>
<input type=button class="button"   value="�Ʊ�λFormatTabs"  onclick="frm.WebOffice.WebDialogs(179);"><br>
<input type=button class="button"   value="��ʽFormatStyle"  onclick="frm.WebOffice.WebDialogs(180);"><br>
<input type=button class="button"   value="�����ʽ�Ի���FormatDefineStyleFont"  onclick="frm.WebOffice.WebDialogs(181);"><br>
<input type=button class="button"   value="�Ʊ�λFormatDefineStyleTabs"  onclick="frm.WebOffice.WebDialogs(183);"><br>
<input type=button class="button"   value="��ʽFormatDefineStyleFrame"  onclick="frm.WebOffice.WebDialogs(184);"><br>
<input type=button class="button"   value="�����ʽ�Ի���FormatParagraph"  onclick="frm.WebOffice.WebDialogs(175);"><br>
<input type=button class="button"   value="�ڵ���ʽFormatSectionLayout"  onclick="frm.WebOffice.WebDialogs(176);"><br>
<input type=button class="button"   value="����FormatColumns"  onclick="frm.WebOffice.WebDialogs(177);"><br>

</td>
<td  nowrap bgcolor=white  valign=top width=80 align=center>
[<b style="color:green">ģ���׺�</b>]<br>
<input type=button class="button"   name="Revision"  value="�ĵ������׺�"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',1);"><br>
<input type=button class="button"   name="Revision"  value="��ǰλ���׺�"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',0);"><br>
<input type=button class="button"   name="Revision"  value="�ĵ��ײ��׺�"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',2);"><br>
[<b style="color:green">�ۼ�ʾ��</b>]<br>
<input type=button class="button" value="���õ�ǰ�û�" onclick="document.all.WebOffice.SetCurrUserName('jenny');alert('�����õ�ǰ�û�Ϊjenny');"><br>
<input type=button class="button" value="��ʼ�޶�SetTrackRevisions" onclick="document.all.WebOffice.SetTrackRevisions(1);"><br>
<input type=button class="button" value="�����޶�SetTrackRevisions" onclick="document.all.WebOffice.SetTrackRevisions(4);"><br>
<input type=button class="button" value="�����޶�ShowRevisions" onclick="document.all.WebOffice.ShowRevisions(0);"><br>
<input type=button class="button" value="�޶�ͳ��" onclick="document.all.WebOffice.showdialog(6);"><br>
<input type=button class="button"   value="ͻ����ʾ�޶�ToolsRevisions"  onclick="frm.WebOffice.WebDialogs(197);"><br>
<input type=button class="button" name="Revision" value="���غۼ�"  onclick="ShowRevision(false)"><br>
<input type=button class="button" value="��ȡ�ۼ�" onclick="ShowRevision(true)"><br>
<input type=button class="button" value="���������޶�" onclick="WebAcceptAllRevisions()"><br>
[<b style="color:green">VBAʾ��</b>]<br>
<select id="template"><option value="1">�����׺�<option value="2">�����׺�2<option value="3">�����׺�<option value="4">�����׺�</select><br>
<input type=button class="button"   value="VBAѡ���׺�"  onclick="WebAddTemplate(template.value)"><br>
<input type=button class="button"   value="VBA�׺춨��"  onclick="WebTempFile();"><br>
<input type=button class="button"   value="���뱾��ͼƬ"  onclick="WebOpenPicture()"><br>
<input type=button class="button"   value="����URLͼƬ"  onclick="WebInsertImage()"><br>
<input type=button class="button" value="�滻Word����" onclick="alert('�����ܽ���ʾ�ؼ���WORD�����еġ��ġ��滻�ɡ�����,ÿִ��һ���滻һ��');document.all.WebOffice.ReplaceText('��','��',1);"><br>
<input type=button class="button"   value="ȡWord����"   onclick="WebGetWordContent()"><br>
<input type=button class="button"   value="дWord����"  onclick="WebSetWordContent()"><br>
<select id="img" style="width:80px"><option value="1">ͼ��1<option value="2">ͼ��2<option value="3">ͼ��3</select><br>

<input type=button class="button"   value="ѡURLͼƬ����"  onclick="WebInsertURLImage(img.value)"><br>
<input type=button class="button"   value="��Excel���"  onclick="WebGetExcelContent()"><br>
<input type=button class="button"   value="�����������"  onclick="WebInsertAfter()"><br>[<b style="color:green">Word�Ի���</b>]<br>
<input type=button class="button"   value="��굽�п�ʼ��"  onclick="frm.WebOffice.WebDialogs(4012);"><br>
<input type=button class="button"   value="��굽��β"  onclick="frm.WebOffice.WebDialogs(4013);"><br>
<input type=button class="button"   value="��굽���ڿ�ʼ��"  onclick="frm.WebOffice.WebDialogs(4014);"><br>
<input type=button class="button"   value="��굽���ں���"  onclick="frm.WebOffice.WebDialogs(4015);"><br>
<input type=button class="button"   value="��굽�ĵ���ʼ��"  onclick="frm.WebOffice.WebDialogs(4016);"><br>
<input type=button class="button"   value="��굽�ĵ�����"  onclick="frm.WebOffice.WebDialogs(4017);"><br>
<input type=button class="button"   value="����EditFind"  onclick="frm.WebOffice.WebDialogs(112);"><br>
<input type=button class="button"   value="�����滻EditReplace"  onclick="frm.WebOffice.WebDialogs(117);"><br>
<input type=button class="button"   value="ҳüNormalViewHeaderArea"  onclick="frm.WebOffice.WebDialogs(155);"><br>
<input type=button class="button"   value="�������ֿ�InsertFrame"  onclick="frm.WebOffice.WebDialogs(158);"><br>
<input type=button class="button"   value="����ָ���InsertBreak"  onclick="frm.WebOffice.WebDialogs(159);"><br>

<input type=button class="button"   value="������עInsertAnnotation"  onclick="frm.WebOffice.WebDialogs(161);"><br>
<input type=button class="button"   value="�����������InsertSymbol"  onclick="frm.WebOffice.WebDialogs(162);"><br>
<input type=button class="button"   value="����ͼƬInsertPicture"  onclick="frm.WebOffice.WebDialogs(163);"><br>
<input type=button class="button"   value="�����ļ�InsertFile"  onclick="frm.WebOffice.WebDialogs(164);"><br>
<input type=button class="button"   value="��������ʱ��InsertDateTime"  onclick="frm.WebOffice.WebDialogs(165);"><br>
<input type=button class="button"   value="������InsertField"  onclick="frm.WebOffice.WebDialogs(166);"><br>
<input type=button class="button"   value="�༭��ǩEditBookmark"  onclick="frm.WebOffice.WebDialogs(168);"><br>
<input type=button class="button"   value="���������MarkIndexEntry"  onclick="frm.WebOffice.WebDialogs(169);"><br>
<input type=button class="button"   value="��������InsertIndex"  onclick="frm.WebOffice.WebDialogs(170);"><br>
<input type=button class="button"   value="����Ŀ¼InsertTableOfContents"  onclick="frm.WebOffice.WebDialogs(171);"><br>
<input type=button class="button"   value="�������InsertObject"  onclick="frm.WebOffice.WebDialogs(172);"><br>

</td>
<td nowrap bgcolor=white valign=top style="padding-top:10px;">
<center><b style="font-family:����;font-size:25px;color:red;"><%=strfname%></b><br>Office�ؼ�������ʾ�ļ���ţ�<input type=text class="text" name=num value="<%=strnum%>">&nbsp;&nbsp;�ļ�����<input type=text class="text" style="width:280" name=fname value="<%=strfname%>">������<input type="file" name="file1" id="file1" />&nbsp;<input type=hidden name=flsid value="<%=strflsid%>"><div align=left>&nbsp;�޷������ؼ�?&nbsp;<a href="http://www.officectrl.com/weboffice/weboffice.rar"><img src="images/weboffice-install.jpg" border=0></a>��ɺ����´򿪱�ҳ�档</div><script language=javascript src="cssjs/webofficeocx.js" charset="utf-8"></script><div id=divinfo style="display:none;">
<iframe style="position:absolute;left:300px;top:250px;background-color:transparent;" id=webifrm src="about:blank" width="580" height="19" scrolling="no" frameborder="0"></iframe>
<div style="position:absolute;left:300px;top:250px;background-color:red;color:white;display:none;" id=webinfo>���ڴ���������ϴ����ݣ�Ϊȷ���ɹ������Ȳ�Ҫִ������������ �����Ե� ������</div></div>

</td></tr></table>&nbsp;&nbsp;<font color=red>&copy;</font>2001-2016 All Rights Reserved!&nbsp;&nbsp;��վ��ַ��<a href="http://www.officectrl.com/">http://www.officectrl.com</a><br></div><br><br></form>
</body></html>

