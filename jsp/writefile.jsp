<%@ page contentType="text/html;charset=gb2312" language="java" %>
<%@ page import="java.io.*" %>
<%@ page import="java.text.*" %>
<%@ page import="java.util.*" %>
<%@ page import="java.sql.*" %>
<%

 
 
	String id = request.getParameter("id");
	String rootpath=application.getRealPath("/");
	String ConnStr="jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ="+rootpath+"\\database\\weboffice.mdb";			
	String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 
	//String ConnStr = "jdbc:odbc:webofficedsn";
	Connection conn = null; 
	ResultSet rs = null; 
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
		Statement stmt = conn.createStatement();
		rs = stmt.executeQuery("select * from dt_document where o_pkid="+id);

			if(rs.next())
			{
			  try
			  { 
				String filename = rs.getString("o_name");
				 
				String flsid = rs.getString("o_flsid");
				String flag = rs.getString("o_flag");
				String filepath="";
				if   (flag.equals("1") || flag.equals("11")) 
				{
					filepath = "\\upload\\"+ flsid +".doc";
				}
				if  (flag.equals("2") || flag.equals("12")) 
				{
					filepath = "\\upload\\"+ flsid +".xls";
				}
				if  (flag.equals("3") || flag.equals("13")) 
				{
					filepath = "\\upload\\"+ flsid +".ppt";
				}
 
				if (filename==null)filename="";
				response.reset();
				response.setContentType("application/octet-stream");
				response.setHeader("Content-Disposition", "attachment; filename="+filename);

				 

				java.io.File file = new java.io.File(rootpath +filepath); 
				java.io.InputStream filedata = new java.io.FileInputStream(file);

				//java.io.InputStream filedata = rs.getBinaryStream("o_file");				
				java.io.OutputStream outStream = response.getOutputStream();
				byte[] buf = new byte[1024];
				int i = 0;
				while((i = filedata.read(buf)) != -1)
				outStream.write(buf, 0, i);
				filedata.close();	
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
		System.err.println("Write QueryErr: " + e.getMessage()); 
	}

	    
		  return;

%>