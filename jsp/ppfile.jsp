<%@ page contentType="text/html;charset=UTF-8" %>
<%@ page import="java.io.*" %>
<%@ page import="java.text.*" %>
<%@ page import="java.util.*" %>
<%@ page import="java.sql.*" %>
<%
//request.setCharacterEncoding("gb2312"); 
request.setCharacterEncoding("UTF-8");
String flsid = request.getParameter("flsid");
String strfile = request.getParameter("file");
String strfileextend = request.getParameter("fileextend");
String strfilenew = request.getParameter("filenew");
//String strfilenew = new String(request.getParameter("filenew").getBytes("iso-8859-1"),"UTF-8");
strfilenew= new String(strfilenew.getBytes("iso-8859-1"),"UTF-8");
 
strfilenew = strfilenew.substring(0,strfilenew.indexOf("."))+".jpg";
strfilenew = strfilenew.replace("幻灯片","a");
// strfilenew = strfilenew.subString()
String rootpath=application.getRealPath("/");
if (strfile==null)strfile="";
if (strfilenew==null)strfilenew="";
if (strfileextend==null)strfileextend="";
//strfile= "";
out.println("1.=====" + strfilenew); 
System.out.println("1.=====" + strfilenew); 

String strSavePath="";
String strImagefloder = rootpath + "\\jsp\\html\\" + strfile + ".files";
//创建图片存储文件夹 
if (!(new java.io.File(strImagefloder).isDirectory())) //如果文件夹不存在
{
	new java.io.File(strImagefloder).mkdir();  
}
System.out.println("2.======="+strImagefloder);  
strSavePath=rootpath + "\\jsp\\html\\" + strfile + ".files\\" + strfilenew ;
System.out.println("3.======="+strSavePath); 

 
FileOutputStream fileOut=new FileOutputStream(strSavePath); 
DataInputStream din=new DataInputStream(request.getInputStream()); 
int formDataLength=request.getContentLength(); 
byte dataBytes[]=new byte[formDataLength];
 //int num=din.skipBytes(10);
din.readFully(dataBytes,0,formDataLength);
fileOut.write(dataBytes); 
fileOut.close(); 
din.close(); 
 
strSavePath=rootpath + "\\jsp\\html\\" + strfile + ".html";
String strValue="<!DOCTYPE html><html><head><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge,chrome=1\"><meta http-equiv=\"content-type\" content=\"text/html;charset=gb2312\"><title></title><body>";
	strValue= strValue + "<div align=center>";
File file = new File(strSavePath);
if(file.exists())
{ 
	file.delete();	
}
file.createNewFile();
 
String strFindDir=rootpath + "\\jsp\\html\\" + strfile + ".files\\";
String strimghtml="";
int b=1;
File f = new File (strFindDir);
File[] files=f.listFiles(); 	//	String nameArray[] = f.list(); 
	for(int i=0;i<files.length;i++){
		//System.out.println("files:" + files[i].getName());

	 	strimghtml = strimghtml + "<img border=1 vspace=8 src="+ strfile + ".files/a"  + b + ".jpg><br><b>" + b + "</b><br><br>"; 
		b=b+1;
}
 strValue = strValue + strimghtml;
//在文件中写入内容
FileOutputStream fos = new FileOutputStream(strSavePath);
fos.write(strValue.getBytes());
fos.flush();
fos.close();
 	

System.out.println("4.======="+strSavePath); 


 

String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 
String ConnStr="jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ="+rootpath+"\\database\\weboffice.mdb";			
//String ConnStr = "jdbc:odbc:webofficedsn";
Connection conn = null;
try 
{
	Class.forName(DBDriver).newInstance();  
} 
catch(java.lang.ClassNotFoundException e)
{ 
	System.err.println("Conn(): " + e.getMessage()); 
}

conn = DriverManager.getConnection(ConnStr); 

String sql=null;
PreparedStatement ps = null;

sql="Update dt_document set o_html=? where o_flsid='"+flsid+"'";
ps = conn.prepareStatement(sql);
ps.setInt(1,1);		 
ps.executeUpdate();

System.out.println(strfilenew); 
%>
