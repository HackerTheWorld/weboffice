<%@ page contentType="text/html;charset=gb2312" %>
<%@ page import="java.io.*" %>
<%@ page import="java.text.*" %>
<%@ page import="java.util.*" %>
<%@ page import="java.sql.*" %>
<%

String oper = request.getParameter("oper");
String flsid = request.getParameter("flsid");
String flag = request.getParameter("flag");
String id = request.getParameter("id");
if (flsid==null)flsid="nothing";
String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 
String rootpath=application.getRealPath("/");
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
try
{ 

		String filepath="\\jsp\\pdf\\"+flsid+".pdf";
		 
 
		filepath = rootpath +filepath;
 
		conn = DriverManager.getConnection(ConnStr); 
		FileOutputStream fileOut=new FileOutputStream(filepath); 
		DataInputStream din=new DataInputStream(request.getInputStream()); 
		int formDataLength=request.getContentLength(); 
		byte dataBytes[]=new byte[formDataLength];
		 //int num=din.skipBytes(10);
		din.readFully(dataBytes,0,formDataLength);
		fileOut.write(dataBytes); 
		fileOut.close(); 
		din.close(); 
		
  
		String sql=null;
		PreparedStatement ps = null;
		 
			sql="Update dt_document set o_pdf=? where o_flsid='"+flsid+"'";
			ps = conn.prepareStatement(sql);
			ps.setInt(1,1);		 
			ps.executeUpdate();
		 
	 
} 
catch(SQLException e)
{ 
	System.err.println("Upload QueryErr: " + e.getMessage()); 
}
conn.close();



%>
