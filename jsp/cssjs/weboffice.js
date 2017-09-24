<!--
function WebOpen()
{
 obj = document.all.item("WebOffice");
 if (obj !='')
 	{
 		//等待控件初始化完毕，时间长短可以根据网络速度设定。
 		setTimeout('openfile()',100);
	}
}
function openfile()
{ 
	try{
		switch(flag)
			{
			case '11':
					frm.WebOffice.Open(strOpenUrl,true,"Word.Document","","");break;
			case '12':
					frm.WebOffice.Open(strOpenUrl,true,"Excel.Sheet","","");break;
			case '13': 
					frm.WebOffice.Open(strOpenUrl,true,"PowerPoint.Show","","");break;
			case '2':
					frm.WebOffice.CreateNew("Excel.Sheet.8");				break;
			case '3':
					frm.WebOffice.CreateNew("PowerPoint.Show");break;
			default: 
					frm.WebOffice.CreateNew("Word.Document");	}
	}
	catch(e)
	{
		//如果发生错误，一般是客户机没有安装WebOffice控件引起的，提示用户下载安装
		//	if(confirm('发生问题，如果您是首次使用，请先下载安装WebOffice控件；\n如果您已经安装过WebOffice控件，请注销Windows重新登陆再试用就可以了。')) 
		//		{
				//window.location.href='http://www.officectrl.com/weboffice/weboffice.rar';
		//		}
	}

}
function WebSave()
{
	alert('本功能将演示Web表单数据、附件与WebOffice文档内容以异步(分离)方式进行保存。\n首先保存WebOffice文档数据，其次再保存Web表单数据。完成后本网页将提示关闭');
	try
	{
		frm.WebOffice.Save(strURL);
		 
	}
	catch(e)
	{
		alert('远程保存Office文档发生错误！');
	}
	return true;
}
function MyTimer()
{ 
	if (autoSave==1)
	{
		info.innerHTML='<br><b style=color:red>保存文档中...</b><br>';
		info.style.display='';
		//frm.WebOffice.Save(strURL);
		frm.WebOffice.HttpInit();
		frm.WebOffice.HttpAddPostString('num',frm.num.value);
		frm.WebOffice.HttpAddPostString('fname',frm.fname.value);
		frm.WebOffice.HttpAddPostString('oper',frm.oper.value);
		frm.WebOffice.HttpAddPostString('flsid',frm.flsid.value);
		frm.WebOffice.HttpAddPostString('flag',frm.flag.value);	
		frm.WebOffice.HttpAddPostFile('file1',frm.file1.value);
		frm.WebOffice.HttpAddPostCurrFile("docfile","");
		frm.WebOffice.HttpPost(strRoot + '/upload.jsp');
		info.innerHTML='<br><b style=color:red>保存完成...</b><br>';
		var timer=setTimeout("MyTimer()",30000);
	}
}
function WebHttpSave()
{
	try
	{
		alert('本功能将演示Web表单数据、附件与WebOffice文档内容同时提交至指定的URL路径接收并存储！');
		frm.WebOffice.HttpInit();
		frm.WebOffice.HttpAddPostString('num',frm.num.value);
		frm.WebOffice.HttpAddPostString('fname',frm.fname.value);
		frm.WebOffice.HttpAddPostString('oper',frm.oper.value);
		frm.WebOffice.HttpAddPostString('flsid',frm.flsid.value);
		frm.WebOffice.HttpAddPostString('flag',frm.flag.value);	 		 
		frm.WebOffice.HttpAddPostFile('file1',frm.file1.value);
		frm.WebOffice.HttpAddPostCurrFile("docfile","");	
		frm.WebOffice.HttpPost(strRoot + '/upload.jsp');
		
		alert('表单数据和office文档保存成功!');
	}
	catch(e)
	{
		alert('远程保存Office文档发生错误！');
	}
 
	return true;

}
function WebSaveRemotePdf()
{
	try
	{
		if(pfile!='' && strPdfSaveUrl!=''){
			if(confirm('注意：本功能需要安装OFFICE2010或以上版本支持。\n如果已安装则可以按确定后开始本次转换操作，\n否则请按取消按钮停止本次操作......')) 
			{
				alert('现在开始执行转换PDF文件的操作，此过程时间可能比较长，\n 请耐心等待转换成功的提示后，再执行其它操作，');
				//首先在本地缓冲区生成PDF文件
				tempPath=frm.WebOffice.TempFilePath;
				var pdfile = tempPath +pfile+'.pdf';
				//控件正在执行任务，让用户暂停其它操作的提示
				webinfo.style.display='';
				divinfo.style.display='';
				//PDF生成方式一：
				frm.WebOffice.ActiveDocument.saveas(pdfile,17);
				//PDF生成方式二：
				//frm.WebOffice.ActiveDocument.ExportAsFixedFormat(pdfile,17);
				//其次上传 pdf文件到服务器，后清空缓冲区中的文档 
				frm.WebOffice.WebSaveAsPDF(pdfile,strPdfSaveUrl);
				//关闭用户暂停其它操作的提示
				webinfo.style.display='none';
				divinfo.style.display='none';

				if(confirm('已将当前打开的WORD文档转成PDF文件并远程保存至服务器成功！\n如是你本地电脑已安装PDF阅读器，现在就可以打开查看，是否现在打开查看？'))
				{
					window.open(strRoot+'pdf/' +pfile+'.pdf','_blank');
				}
				
			}	
		}
	}
	catch(e)
	{
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('可能原因有：1、您本地的Office版本过低不支持将WORD文件转为PDF文件,请安装OFFICE2010以上版本！ \n            2、远程保存时服务器端发生错误！  ');
	}
}
function WebSaveRemoteHTML()
{
	try
	{
		if(pfile!='' && strHTMLSaveUrl!=''){
			alert('注意：现在开始执行转换HTML文件的操作，此过程时间可能比较长，\n 请耐心等待转换成功的提示后，再执行其它操作，\n现在请按确定后开始本次转换操作......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend =".html";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//控件正在执行任务，让用户暂停其它操作的提示
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.saveas(htmlfullpath,8) ;
			openfile(); 	 
			frm.WebOffice.WebSaveAsHTML(htmlpath,htmlname,htmlExtend,strHTMLSaveUrl);
			//关闭用户暂停其它操作的提示
			webinfo.style.display='none';
			divinfo.style.display='none';
			if(confirm('已将当前打开的WORD文档转成HTML文件并远程保存至服务器成功！是否现在打开查看？'))
			{
				window.open(strRoot+'html/' +pfile+htmlExtend,'_blank');
			} 
		}
	}
	catch(e)
	{
		alert('请注销您本地电脑后重试，如果仍有问题，请系管理员！');
	}	
}
function WebSaveLocalPdf()
{
	try
	{
		if(pfile!=''){
			var pdfile = 'C:\\'+pfile+'.pdf';
			//控件正在执行任务，让用户暂停其它操作的提示
			divinfo.style.display='';
			webinfo.style.display='';
			//PDF生成方式一：
			//frm.WebOffice.ActiveDocument.saveas(pdfile,17);
			//PDF生成方式二：
			frm.WebOffice.ActiveDocument.ExportAsFixedFormat(pdfile,17);
			//关闭用户暂停其它操作的提示
			webinfo.style.display='none';
			divinfo.style.display='none';
			alert('已在C盘根目录下生成'+pdfile+'，请到你本地电脑的C盘目录查看！');
		}
	}
	catch(e)
	{
		alert('您本地的Office版本过低不支持将WORD转为PDF,请安装OFFICE2010以上版本！ ');
	}
}
function WebSaveLocalHTML()
{ 
	try
	{
		if(pfile!=''){
			var htmlpath="C:\\"
			var htmlname=pfile;
			var htmlExtend =".html";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//控件正在执行任务，让用户暂停其它操作的提示
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.saveas(htmlfullpath,8);
			//关闭用户暂停其它操作的提示
			webinfo.style.display='none';
			divinfo.style.display='none';
			//openfile();
			alert('已在C盘根目录下生成'+htmlfullpath+'，请到你本地电脑的C盘目录查看！');
		}
	}
	catch(e)
	{
		alert('请注销您本地电脑后重试，如果仍有问题，请系管理员！');
	}
}
function WebSaveXLSAsPDF()
{ 
	try{ 
	
		if(pfile!='' && strPdfSaveUrl!=''){
				if(confirm('注意：本功能需要安装OFFICE2010或以上版本支持。\n如果已安装则可以按确定后开始本次转换操作，\n否则请按取消按钮停止本次操作......')) 
				{
					alert('现在开始执行转换PDF文件的操作，此过程时间可能比较长，\n 请耐心等待转换成功的提示后，再执行其它操作，');
					//首先在本地缓冲区生成PDF文件
					tempPath=frm.WebOffice.TempFilePath;
					var pdfile = tempPath +pfile+'.pdf';
					//控件正在执行任务，让用户暂停其它操作的提示
					divinfo.style.display='';
					webinfo.style.display='';					 
					//PDF生成方式一： 
					frm.WebOffice.ActiveDocument.Application.ActiveWindow.ActiveSheet.ExportAsFixedFormat(0,pdfile);
					//其次上传 pdf文件到服务器，后清空缓冲区中的文档 
					frm.WebOffice.WebSaveAsPDF(pdfile,strPdfSaveUrl);
					//关闭用户暂停其它操作的提示
					webinfo.style.display='none';
					divinfo.style.display='none';
					if(confirm('已将当前打开的EXCEL文档转成PDF文件并远程保存至服务器成功！\n如是你本地电脑已安装PDF阅读器，现在就可以打开查看，是否现在打开查看？'))
					{
						window.open(strRoot+'pdf/' +pfile+'.pdf','_blank');
					}
				}
			} 
	
	}
	catch(e)
	{
		alert('您本地的Office版本过低不支持将EXCEL转为PDF,请安装OFFICE2010以上版本！ ');
	}
}
function WebSaveXLSLocalPDF(){
	try
	{
		var pdfile = 'C:\\'+pfile+'.pdf';
		//控件正在执行任务，让用户暂停其它操作的提示
		divinfo.style.display='';
		webinfo.style.display='';
		frm.WebOffice.ActiveDocument.Application.ActiveWindow.ActiveSheet.ExportAsFixedFormat(0,pdfile);
		//关闭用户暂停其它操作的提示
		webinfo.style.display='none';
		divinfo.style.display='none'; 
		alert('已在C盘根目录下生成'+pdfile+'，请到你本地电脑的C盘目录查看！');
	}
	catch(e)
	{
		 
	}
}

function WebSaveXLSLocalHTML(){
	try
	{
		var pdfile = 'C:\\'+pfile+'.html';
		//控件正在执行任务，让用户暂停其它操作的提示
		divinfo.style.display='';
		webinfo.style.display='';
		frm.WebOffice.ActiveDocument.Application.ActiveWorkbook.SaveAs(pdfile,44); 
		//关闭用户暂停其它操作的提示
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('已在C盘根目录下生成'+pdfile+'，请到你本地电脑的C盘目录查看！');
	}
	catch(e)
	{
		
	}
}
function WebSaveXLSAsHTML(){	
		try
	{
		if(pfile!='' && strHTMLSaveUrl!=''){
			alert('注意：现在开始执行转换HTML文件的操作，此过程时间可能比较长，\n 请耐心等待转换成功的提示后，再执行其它操作，\n现在请按确定后开始本次转换操作......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend =".html";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//控件正在执行任务，让用户暂停其它操作的提示
			//divinfo.style.display='';
			//webinfo.style.display='';			
			frm.WebOffice.ActiveDocument.Application.ActiveWorkbook.SaveAs(htmlfullpath,44);	
			
			frm.field1.value=htmlpath;
			frm.field2.value=htmlname;
			frm.field3.value=htmlExtend;
			frm.field4.value=strHTMLSaveUrl;
			 
			frm.action="excelhtml.jsp";
			frm.submit();
			//关闭用户暂停其它操作的提示
		//	webinfo.style.display='none';
		//	divinfo.style.display='none';
	 
		
		}
	}
	catch(e)
	{
		alert(e);
		alert('请注销您本地电脑后重试，如果仍有问题，请系管理员！');
	}	

}
function WebSaveXLSAsMHT(){
	try
	{
		if(pfile!='' && strHTMLSaveUrl!=''){
			alert('注意：现在开始执行转换HTML文件的操作，此过程时间可能比较长，\n 请耐心等待转换成功的提示后，再执行其它操作，\n现在请按确定后开始本次转换操作......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend =".mht";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//控件正在执行任务，让用户暂停其它操作的提示
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.Application.ActiveWorkbook.SaveAs(htmlfullpath,45);			 
			frm.action="excelhtml.jsp";
			frm.field1.value=htmlpath;
			frm.field2.value=htmlname;
			frm.field3.value=htmlExtend;
			frm.field4.value=strHTMLSaveUrl;
			frm.submit(); 
			//关闭用户暂停其它操作的提示
			webinfo.style.display='none';
			divinfo.style.display='none';
		}
	}
	catch(e)
	{
		alert('请注销您本地电脑后重试，如果仍有问题，请系管理员！');
	}	
}


function WebSavePPTLocalPDF(){
	try
	{	
		//控件正在执行任务，让用户暂停其它操作的提示
		divinfo.style.display='';
		webinfo.style.display='';
		var pdfile = 'C:\\'+pfile+'.pdf';
		frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (pdfile,32);
		//关闭用户暂停其它操作的提示
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('已在C盘根目录下生成'+pdfile+'，请到你本地电脑的C盘目录查看！');
	}
	catch(e)
	{
		 alert(e);
	}
}

function WebSavePPTLocalJPG(){
	try
	{
		//控件正在执行任务，让用户暂停其它操作的提示
		divinfo.style.display='';
		webinfo.style.display='';
		var pdfile = 'C:\\'+pfile;
		frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (pdfile,17);//jpg
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",6);//rtf
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",7);//pps
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",16);//gif
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",17);//jpg
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",18);//png
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",19);//bmp
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",21);//tif word
		//关闭用户暂停其它操作的提示
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('已把PPT转成图片放在'+pdfile+'目录下，请到你本地电脑的C盘目录查看！');
	}
	catch(e)
	{
		alert(e);
	}
}
function WebSavePPTAsPDF()
{ 
	try{ 
	
		if(pfile!='' && strPdfSaveUrl!=''){
				if(confirm('注意：本功能需要安装OFFICE2010或以上版本支持。\n如果已安装则可以按确定后开始本次转换操作，\n否则请按取消按钮停止本次操作......')) 
				{
					alert('现在开始执行转换PDF文件的操作，此过程时间可能比较长，\n 请耐心等待转换成功的提示后，再执行其它操作，');
					//控件正在执行任务，让用户暂停其它操作的提示
					divinfo.style.display='';
					webinfo.style.display='';
					//首先在本地缓冲区生成PDF文件
					tempPath=frm.WebOffice.TempFilePath;
					var pdfile = tempPath +pfile+'.pdf';
					//PDF生成方式一： 
					frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (pdfile,32);
					//其次上传 pdf文件到服务器，后清空缓冲区中的文档 
					frm.WebOffice.WebSaveAsPDF(pdfile,strPdfSaveUrl);
					//关闭用户暂停其它操作的提示
					webinfo.style.display='none';
					divinfo.style.display='none';
					if(confirm('已将当前打开的PPT文档转成PDF文件并远程保存至服务器成功！\n如是你本地电脑已安装PDF阅读器，现在就可以打开查看，是否现在打开查看？'))
					{
						window.open(strRoot+'pdf/' +pfile+'.pdf','_blank');
					}
				}
			} 
	
	}
	catch(e)
	{
		alert('您本地的Office版本过低不支持将PPT转为PDF,请安装OFFICE2010以上版本！ ');
	}
}
function WebSavePPTAsHTML(){	
	try
	{
		if(pfile!='' && strppFileSaveUrl!=''){
			alert('注意：现在开始执行转换HTML文件的操作，此过程时间可能比较长，\n 请耐心等待转换成功的提示后，再执行其它操作，\n现在请按确定后开始本次转换操作......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend ='';
			var htmlfullpath= htmlpath+pfile;			
			//控件正在执行任务，让用户暂停其它操作的提示
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (htmlfullpath,17);//jpg 
	 
			frm.WebOffice.WebSaveFormFolder(htmlfullpath+'\\',strppFileSaveUrl+'&file='+pfile);	
		 
			//关闭用户暂停其它操作的提示
			webinfo.style.display='none';
			divinfo.style.display='none';
			if(confirm('已将当前打开的PPT文档转成HTML文件并远程保存至服务器成功！\n现在就可以打开查看，是否现在打开查看？'))
					{
						window.open(strRoot+'html/' +pfile+'.html','_blank');
					}
		}
	}
	catch(e)
	{
		alert('请注销您本地电脑后重试，如果仍有问题，请系管理员！');
	}	

}
function WebSaveLocal()
{
	//弹出保存对话框
	frm.WebOffice.showdialog(3);
}
function WebOpenLocal()
{
	//弹出打开对话框
	frm.WebOffice.showdialog(1);
}

function WebDocReload()
{
	//alert('本功能将重新装载本网页的所有内容，即刷新！');
	location.reload();	
}
function WebOpenPicture()
{
	//弹出插入对话框
	frm.WebOffice.ActiveDocument.Application.Dialogs(163).Show();
}
 
function WebDocPageSetup()
{
	frm.WebOffice.showdialog(5);
}
function ShowRevision(boolvalue)
{	
	frm.WebOffice.ActiveDocument.ShowRevisions = boolvalue;
}
function WebAcceptAllRevisions()
{
	frm.WebOffice.ActiveDocument.AcceptAllRevisions();
}
function WebSignature(str)
{
	 
	var strPic ='';
	switch(str)
	{
		//此处可以是完整的URL
	case '1':
	strPic = strRoot + "/images/001.gif";
	break;
	case '2':
	strPic = strRoot + "/images/002.gif";
	break;
	case '3':
	strPic = strRoot + "/images/003.gif";
	break;		
	}  

	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1',strPic,'::FLOATJPG::');
	var doc = frm.WebOffice.ActiveDocument;
	//doc.Shapes.AddPicture(strPic, false, true,100,0,207,209,doc.Application.Selection.Range);
	doc.Shapes(doc.Shapes.Count).Select(); 
	var range = doc.Application.Selection.ShapeRange;
	range.WrapFormat.Type = 3;
	range.PictureFormat.TransparentBackground = true;
	range.PictureFormat.TransparencyColor = 0xFFFFFF;
	range.Fill.Visible = false;

}


function WebAddFloatPic()
{
 
	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/images/weboffice.jpg','::FLOATJPG::');
	 
} 
function WebAddPic()
{
	
	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/images/weboffice.jpg','::JPG::');
	 
}
function WebDocSignature()
{
	try{
	
	frm.WebOffice.WebSign();	 
	var doc = frm.WebOffice.ActiveDocument;	
	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1','c:\\Sign.bmp','::FLOATJPG::');
	//doc.Shapes.AddPicture('c:\\Sign.bmp', false, true,100,0,219,112,doc.Application.Selection.Range);
	doc.Shapes(doc.Shapes.Count).Select(); 
	var range = doc.Application.Selection.ShapeRange;
	range.WrapFormat.Type = 3;
	range.PictureFormat.TransparentBackground = true;
	range.PictureFormat.TransparencyColor = 0xFFFFFF;
	range.Fill.Visible = false;
	//frm.WebOffice.WebSignTempFileDel(); 
	//var strFile = frm.WebOffice.WebSignTempFile;	
	  
	}
	catch(E)
	{
		
	}
}
function WebTempFile(str)
{
	var strValue='';
	switch(str)
	{
		case '1':
		strValue='OfficeCTRL技术开发中心发文';
			break;
		case '2':
		strValue='OfficeCTRL技术开发中心公文';
		var doc = frm.WebOffice.ActiveDocument;	
		doc.Shapes.AddPicture(strRoot + "/images/weboffice.jpg",false, true,0,-60);
			break;
		case '3':
		strValue='OfficeCTRL技术开发中心公文';		
			break;
		case '4':
		strValue='OfficeCTRL技术开发中心收文';
			break;
		default:
		strValue='电子政务文件';
	}
	//画线
	var object=frm.WebOffice.ActiveDocument;
	//var myl=object.Shapes.AddLine(100,60,305,60)
	//myl.Line.ForeColor=255;
	//myl.Line.Weight=2;
	//var myl1=object.Shapes.AddLine(326,60,520,60)
	//myl1.Line.ForeColor=255;
	//myl1.Line.Weight=2;

	//object.Shapes.AddLine(200,200,450,200).Line.ForeColor=6;
   	var myRange=frm.WebOffice.ActiveDocument.Range(0,0);
	myRange.Select();
	var mtext="★";
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mtext+"\n");
   	var myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
   	myRange.ParagraphFormat.LineSpacingRule =1.5;
   	myRange.font.ColorIndex=6;
   	myRange.ParagraphFormat.Alignment=1;
   	myRange=frm.WebOffice.ActiveDocument.Range(0,0);
	myRange.Select();
	mtext="[２０16]１72号";
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mtext+"\n");
	myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
	myRange.ParagraphFormat.LineSpacingRule =1.5;
	myRange.ParagraphFormat.Alignment=1;
	myRange.font.ColorIndex=1;	
	mtext=strValue;
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mtext+"\n");
	myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
	myRange.ParagraphFormat.LineSpacingRule =1.5;	
	//myRange.Select();
	myRange.Font.ColorIndex=6;
	myRange.Font.Name="仿宋_GB2312";
	myRange.font.Bold=true;
	myRange.Font.Size=28;
	myRange.ParagraphFormat.Alignment=1;	
	//myRange=myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
	frm.WebOffice.ActiveDocument.PageSetup.LeftMargin=70;
	frm.WebOffice.ActiveDocument.PageSetup.RightMargin=70;
	frm.WebOffice.ActiveDocument.PageSetup.TopMargin=70;
	frm.WebOffice.ActiveDocument.PageSetup.BottomMargin=70;
}
function WebSetWordTable()
{
	try{
		var mText="",mTmp="",iColumns=10,iCells=10,iPost,iold=-1;
		var myRange=frm.WebOffice.ActiveDocument.Range(0,0);     //光标位置
		frm.WebOffice.ActiveDocument.Tables.Add(myRange,10,10);   //生成表格 
/*
		for (var n=0; n<iColumns; n++)
		{
			for (var i=0; i<iCells; i++)
			{
				iPos  = mText.indexOf(";",1+iold);
				mTmp = mText.substring(iold+1,iPos);
				frm.WebOffice.ActiveDocument.Tables(1).Columns(n+1).Cells(i+1).Range.Text=mTmp;   //填充单元值
				iold = iPos; 
			}
		}   
	*/
	} 
	catch(e)
	{
		alert(e);
	}

}
function WebGetWordContent()
{
  try{
    alert(frm.WebOffice.ActiveDocument.Content.Text);
  }catch(e){}
}
function WebSetWordContent()
{
	var mText=window.prompt("请输入内容:","测试内容");
	if (mText==null){
		return (false);
	}
	else
	{
		 //下面为显示选中的文本
		 //alert(frm.WebOffice.ActiveDocument.Application.Selection.Range.Text);
		 //下面为在当前光标出插入文本
		 frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mText+"\n");
		 //下面为在第一段后插入文本
		 //frm.WebOffice.ActiveDocument.Application.ActiveDocument.Range(1).InsertAfter(mText);
	}
}
function WebGetExcelContent()
{	try{
	frm.WebOffice.ActiveDocument.Application.Sheets(1).Select;
    frm.WebOffice.ActiveDocument.Application.Range("C5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "126";
    frm.WebOffice.ActiveDocument.Application.Range("C6").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "446";
    frm.WebOffice.ActiveDocument.Application.Range("C7").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "556";
    frm.WebOffice.ActiveDocument.Application.Range("C5:C8").Select;
    frm.WebOffice.ActiveDocument.Application.Range("C8").Activate;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)";
    frm.WebOffice.ActiveDocument.Application.Range("D8").Select;
    alert(frm.WebOffice.ActiveDocument.Application.Range("C8").Text);
	 }catch(e){
		alert('此功能对Excel文档有效!');
	 }
}
//作用：保护工作表单元
function WebSheetsLock(){
	try{
    frm.WebOffice.ActiveDocument.Application.Sheets(1).Select;
    frm.WebOffice.ActiveDocument.Application.Range("A1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "产品";
    frm.WebOffice.ActiveDocument.Application.Range("B1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "价格";
    frm.WebOffice.ActiveDocument.Application.Range("C1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "详细说明";
    frm.WebOffice.ActiveDocument.Application.Range("D1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "库存";
    frm.WebOffice.ActiveDocument.Application.Range("A2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "书签";
    frm.WebOffice.ActiveDocument.Application.Range("A3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "毛笔";
    frm.WebOffice.ActiveDocument.Application.Range("A4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "钢笔";
    frm.WebOffice.ActiveDocument.Application.Range("A5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "尺子";
    frm.WebOffice.ActiveDocument.Application.Range("B2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "0.5";
    frm.WebOffice.ActiveDocument.Application.Range("C2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "樱花";
    frm.WebOffice.ActiveDocument.Application.Range("D2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "300";
    frm.WebOffice.ActiveDocument.Application.Range("B3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "2";
    frm.WebOffice.ActiveDocument.Application.Range("C3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "狼毫";
    frm.WebOffice.ActiveDocument.Application.Range("D3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "50";
    frm.WebOffice.ActiveDocument.Application.Range("B4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "3";
    frm.WebOffice.ActiveDocument.Application.Range("C4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "蓝色";
    frm.WebOffice.ActiveDocument.Application.Range("D4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "90";
    frm.WebOffice.ActiveDocument.Application.Range("B5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "1";
    frm.WebOffice.ActiveDocument.Application.Range("C5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "20cm";
    frm.WebOffice.ActiveDocument.Application.Range("D5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "40";
    //保护工作表
    frm.WebOffice.ActiveDocument.Application.Range("B2:D5").Select;
    frm.WebOffice.ActiveDocument.Application.Selection.Locked = false;
    frm.WebOffice.ActiveDocument.Application.Selection.FormulaHidden = false;
    frm.WebOffice.ActiveDocument.Application.ActiveSheet.Protect(true,true,true);
    alert("已经保护工作表，只有B2-D5单元格可以修改。");
	 }catch(e){
		alert('此功能对Excel文档有效!');
	 }
}
//作用：获取文档页数
function WebDocumentPageCount(){
	var intPageTotal;
	intPageTotal = frm.WebOffice.ActiveDocument.Application.ActiveDocument.BuiltInDocumentProperties(14);
	alert("文档页总数："+intPageTotal);
}

function WebTitlebar(boolvalue)
{
	frm.WebOffice.Titlebar = boolvalue;
}
function WebToolbar(boolvalue)
{
	frm.WebOffice.Toolbars = boolvalue;
}
function WebMenubar(boolvalue)
{
	   frm.WebOffice.MenuBars =boolvalue;
	   
}
function WebInsertImage()
{

	frm.WebOffice.ActiveDocument.Application.Selection.InlineShapes.AddPicture(strRoot+"/images/login.gif",false,true);

}
function WebInsertURLImage(str)
{
	var fileName='';

	switch(str)
	{
	case '2':
		fileName = strRoot + "/images/sec.jpg";		
		break;
	case '3':
		fileName = strRoot + "/images/buy.gif";
		break;
	default:
		fileName = strRoot + "/images/180.jpg";
	}
	frm.WebOffice.ActiveDocument.Application.Selection.InlineShapes.AddPicture(fileName,false,true);
}
function WebAddTemplate(str)
{
	WebTempFile(str);	
}
function WebPrintDirc()
{
frm.WebOffice.ActiveDocument.PrintOut();
}
function WebDocPrint()
{
 frm.WebOffice.printout(true);
}
function WebDocPrint2()
{
	 frm.WebOffice.ActiveDocument.Application.Dialogs(88).Show();
}
function WebDocPrintPreView()
{
frm.WebOffice.ActiveDocument.Application.PrintPreview=1;
}
function WebSaveC()
{
	try{
	frm.WebOffice.ActiveDocument.SaveAs("c:\\a.doc");
	alert('已保存到c盘a.doc');
	}
	 catch(e){
		alert('发生错误!');
	 }

}function WebInsertAfter()
{
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter('公司名称');
	//将光标移到本行文字后面
	frm.WebOffice.ActiveDocument.Application.Dialogs(4013).Show();
}
function WebGetAllMark()
{
	var iCount=frm.WebOffice.ActiveDocument.BookMarks.count;
	for (i=1;i<=iCount ; i++ )
	{
		alert(frm.WebOffice.ActiveDocument.BookMarks.item(i).Name);
		 
	}
}
//-->