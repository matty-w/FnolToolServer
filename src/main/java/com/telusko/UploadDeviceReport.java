package com.telusko;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Servlet implementation class UploadDeviceReport
 */
public class UploadDeviceReport extends HttpServlet 
{
	private static final long serialVersionUID = 1L;

	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException 
	{
		try
		{
			File deviceDirectory = new File("C:/Users/Matt/ServerPlayaround/FileUploadDemo/deviceReport/");
			
			FileUtils.cleanDirectory(deviceDirectory);
			
			ServletFileUpload sf = new ServletFileUpload(new DiskFileItemFactory());
			List<FileItem> multiFiles = sf.parseRequest(request);
			
			for(FileItem item : multiFiles)
			{
				System.out.println(item.getName());
				item.write(new File("C:/Users/Matt/ServerPlayaround/FileUploadDemo/deviceReport/" + FilenameUtils.getName(item.getName())));
			}
			
			File[] listFiles = deviceDirectory.listFiles();
			
			File deviceReport = listFiles[0];
			
			convertDeviceReportToXlsx(deviceReport);
			deviceReport.delete();
			
			System.out.println("Done");
			
		}
		catch(Exception e)
		{
			
		}

	}
	
	private void convertDeviceReportToXlsx(File deviceReport) 
	{
		try
		{
			String parent = deviceReport.getParent();
			String fileName = deviceReport.getName();
			int end = fileName.indexOf(".");
			String fileNameCut = fileName.substring(0, end);
			FileInputStream fis = new FileInputStream(deviceReport);
	        BufferedReader reader = new BufferedReader(new InputStreamReader(fis));
	        XSSFWorkbook deviceReportXlsx = new XSSFWorkbook();
	        FileOutputStream writer = new FileOutputStream(new File(parent+"\\"+fileNameCut+".xlsx") );
	        XSSFSheet mySheet = deviceReportXlsx.createSheet();
	        String line= "";
	        int rowNo=0;
	        
	        while ( (line=reader.readLine()) != null )
	        {
	            String[] columns = line.split(",");
	            XSSFRow myRow =mySheet.createRow(rowNo);
	            for (int i=0;i<columns.length;i++)
	            {
	                XSSFCell myCell = myRow.createCell(i);
	                myCell.setCellValue(columns[i]);
	            }
	             rowNo++;
	        }
	        deviceReportXlsx.write(writer);
	        reader.close();
	        writer.close();
	        //deviceReportSaved = deviceReportXlsx;
		}
		catch(Exception e)
		{
		}
	}
	

}
