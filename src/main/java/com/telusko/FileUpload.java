package com.telusko;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * Servlet implementation class FileUpload
 */
public class FileUpload extends HttpServlet
{
	static XSSFWorkbook deviceReportSaved = null;
	List<List<String>> errors = new ArrayList<List<String>>();
	private static final long serialVersionUID = 1L;
	private static final int BYTES_DOWNLOAD = 4096;

	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException 
	{
		try
		{
			String folderDateTimeName = createName();
			String homeSystem = System.getProperty("user.home")+"/ServerPlayaround/FileUploadDemo/UserDataFiles/";
			File emailFolder = new File(homeSystem+"EmailFolder_"+folderDateTimeName);
			File deviceReportFolder = new File(homeSystem+"DeviceFolder_"+folderDateTimeName);
			File claimsFileFolder = new File(homeSystem+"claimsFileFolder_"+folderDateTimeName);
			emailFolder.mkdir();
			deviceReportFolder.mkdir();
			claimsFileFolder.mkdir();
			
			
			
			downloadFileToServer(request, emailFolder, deviceReportFolder, claimsFileFolder);
			
			
			convertDeviceReportToXlsx(deviceReportFolder);
			
			List<String> defaultOptions = new ArrayList<String>();
			defaultOptions.add("Tesco vehicle registration number");
			defaultOptions.add("Device ID");
			defaultOptions.add("");
			defaultOptions.add("Collision date");
			defaultOptions.add("Collision time");
			defaultOptions.add("");
			defaultOptions.add("sopp+sopp Reference number ");
			defaultOptions.add("Collision causation code");
			defaultOptions.add("FNOL File Name");
			
			
			File[] listEmailFiles = emailFolder.listFiles(new FilenameFilter() 
			{
				
				public boolean accept(File dir, String name) 
				{
					return name.endsWith(".msg");
				}
			});
			
			XSSFWorkbook deviceReport = convertDeviceReportToXlsx(deviceReportFolder);
			
			File[] fnols = createFnolFileList(listEmailFiles);
			
			List<List<String>> requiredFnolData = fnolDataForSpreadsheet(fnols, defaultOptions, deviceReport);
			System.out.println(requiredFnolData.size());
			List<String> testing = requiredFnolData.get(0);
			System.out.println(testing.get(0));
			List<List<String>> errorData = getErrorData();
			
			File resultsWorkbook = createResultsFileWithFnolDataCustom(defaultOptions, requiredFnolData, errorData);
			
			reformatCells(resultsWorkbook);
			
			
			downloadFile(response, resultsWorkbook);
			
			deleteFileAndFolder(resultsWorkbook);
			
			cleanRemainingFolders(emailFolder, deviceReportFolder, claimsFileFolder);
			
			
		}
		catch(Exception e)
		{
			System.out.println(e);
		}

	}

	private XSSFWorkbook convertDeviceReportToXlsx(File deviceReportFolder) 
	{
		try
		{
			File[] listFiles = deviceReportFolder.listFiles();
			File f = listFiles[0];
			String fileName = f.getName();
			int end = fileName.indexOf(".");
			String fileNameCut = fileName.substring(0, end);
			FileInputStream fis = new FileInputStream(f);
	        BufferedReader reader = new BufferedReader(new InputStreamReader(fis));
	        XSSFWorkbook deviceReportXlsx = new XSSFWorkbook();
	        FileOutputStream writer = new FileOutputStream(new File(deviceReportFolder.getAbsolutePath()+"\\"+fileNameCut+".xlsx") );
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
	        return deviceReportXlsx;
		}
		catch(Exception e)
		{
			return null;
		}
	}

	private void downloadFileToServer(HttpServletRequest request, File emailFolder, File deviceFolder, File claimsFolder)
			throws FileUploadException, Exception 
	{
		ServletFileUpload drUpload = new ServletFileUpload(new DiskFileItemFactory());
		List<FileItem> deviceReport = drUpload.parseRequest(request);
		
		for(FileItem item : deviceReport)
		{
			System.out.println(item.getName());
			if(item.getName().equals("") || item.getName().equals(null))
				continue;
			else if(item.getName().endsWith(".msg"))
			{
				String emailFolderPath = emailFolder.getAbsolutePath();
				item.write(new File(emailFolderPath +"/"+ FilenameUtils.getName(item.getName())));
			}
			else if(item.getName().endsWith(".csv"))
			{
				String deviceReportPath = deviceFolder.getAbsolutePath();
				item.write(new File(deviceReportPath +"/"+ FilenameUtils.getName(item.getName())));
			}
			else if(item.getName().endsWith(".xlsx"))
			{
				String claimsFilePath = claimsFolder.getAbsolutePath();
				item.write(new File(claimsFilePath +"/"+ FilenameUtils.getName(item.getName())));
			}
				
		}
	}
	
	private void cleanRemainingFolders(File emailFolder, File deviceFolder, File claimsFolder) throws IOException
	{
		FileUtils.deleteDirectory(emailFolder);
		FileUtils.deleteDirectory(deviceFolder);
		FileUtils.deleteDirectory(claimsFolder);
	}
	
	private void deleteFileAndFolder(File file) throws IOException
	{
		String dirPath = file.getParent();
		File dir = new File(dirPath);
		FileUtils.deleteDirectory(dir);
	}
	
	private List<List<String>> getErrorData() 
	{
		List<List<String>> errorData = errors;
		errors.clear();
		return errorData;
	}

	private File[] createFnolFileList(File[] emailFiles) throws Exception
	{
		String f = "C:/Users/Matt/ServerPlayaround/FileUploadDemo";
		List<String> fileNames = new ArrayList<String>();
		
		for(int i = 0; i < emailFiles.length; i++)
		{
			if(emailFiles[i].isFile())
			{
				String name = emailFiles[i].getPath();
				fileNames.add(name);
			}
		}
		
		for(int i = 0; i < fileNames.size(); i++)
		{
			String msgFileString = fileNames.get(i);
			MAPIMessage msg = new MAPIMessage(msgFileString);
			
			AttachmentChunks[] attachments = msg.getAttachmentFiles();
			if(attachments.length > 0) 
			{
	            for (AttachmentChunks a  : attachments) 
	            {
	                ByteArrayInputStream fileIn = new ByteArrayInputStream(a.attachData.getValue());
	                File msgFile = new File(f+"/tempFnols", a.attachLongFileName.toString()); // output
	                OutputStream fileOut = null;
	                try 
	                {
	                    fileOut = new FileOutputStream(msgFile);
	                    byte[] buffer = new byte[2048];
	                    int bNum = fileIn.read(buffer);
	                    while(bNum > 0) 
	                    {
	                        fileOut.write(buffer);
	                        bNum = fileIn.read(buffer);
	                    }
	                }
	                finally 
	                {
	                    try 
	                    {
	                        if(fileIn != null) 
	                        {
	                            fileIn.close();
	                        }
	                    }
	                    finally 
	                    {
	                        if(fileOut != null) 
	                        {
	                            fileOut.close();
	                        }
	                    }
	                }
	            }
	        }
	        else
	        {
	        }
		}
		
		File fnolFolder = new File(f+"/tempFnols");
		File[] listFnols = fnolFolder.listFiles(new FilenameFilter() 
		{
			public boolean accept(File dir, String name) 
			{
				if(name.contains("ref") || name.contains("Ref"))
					return name.endsWith(".xls");
				return false;
			}
		});
		
		return listFnols;
	}
	
	private List<List<String>> fnolDataForSpreadsheet(File[] listOfFnols, List<String> defaultOptions, XSSFWorkbook deviceReportFolder)
	{
		try
		{
			
			List<List<String>> fnolData = new ArrayList<List<String>>();
			List<List<String>> fnolUnreliableList = new ArrayList<List<String>>();
			
			
			
			
			for(int i = 0; i < listOfFnols.length; i++)
			{
				File fnol = listOfFnols[i];
				System.out.println(fnol.getName());
				FileInputStream fis = new FileInputStream(fnol);
				HSSFWorkbook fnolWorkbook = new HSSFWorkbook(fis);
				HSSFSheet sheet = fnolWorkbook.getSheetAt(0);
				
				Row testRow = sheet.getRow(101);
				Cell testCell = testRow.getCell(0);
				String cellValue = testCell.getStringCellValue();
				
				if(!(cellValue.equals("Driver Trainer Email Address")))
				{
					List<String> fnolUnreliable = new ArrayList<String>();
					String reason = "";
					String fnolFileName = fnol.getName();
					if(cellValue.equals(""))
						reason = "FNOL Is Too Short To Accurately Mine For Results. Check File Manually";
					else
						reason = "FNOL Is Too Long To Accurately Mine For Results. Check File Manually";
					fnolUnreliable.add(fnolFileName);
					fnolUnreliable.add(reason);
					fnolUnreliableList.add(fnolUnreliable);
					errors = fnolUnreliableList;
					continue;
				}
				else
				{
					List<String> values = new ArrayList<String>();
					for(String selectedOption : defaultOptions)
					{
						Iterator<Row> rowIterator = sheet.iterator();
						if(selectedOption.equals(""))
						{
							values.add("");
							continue;
						}
						else if(selectedOption.equals("FNOL File Name"))
						{
							String fnolName = fnol.getName();
							values.add(fnolName);
							continue;
						}
						
						while(rowIterator.hasNext())
						{
							//String[] options = option.split(Pattern.quote("||"));
							//int rowNum = Integer.parseInt(options[0]);
							Row row = rowIterator.next();
							//int fnolRowNum = row.getRowNum();
							Cell titleCell = row.getCell(0);
							String cellTitle = titleCell.getStringCellValue();
							

							if(cellTitle.equals(selectedOption))
							{
								System.out.println(cellTitle);
								Cell cell = row.getCell(1);
								cell.setCellType(Cell.CELL_TYPE_STRING);
								String cv = cell.getStringCellValue();
								values.add(cv);
								if(cellTitle.equals("Tesco vehicle registration number"))
								{
									String id = calculateDeviceId(cell, deviceReportFolder);
									values.add(id);
								}
							}
							
						}
					}
					fnolData.add(values);
					
					
					/*List<String> values = new ArrayList<String>();
					String fnolName = fnol.getName();
					values.add(fnolName);
					while(rowIterator.hasNext())
					{
						Row row = rowIterator.next();
						Cell titleCell = row.getCell(0);
						String cellTitle = titleCell.getStringCellValue();
						
						for(int k = 0; k < defaultOptions.size(); k++)
						{
							String[] tempItems = defaultOptions.get(k).split(Pattern.quote("||"));
							int rowNumber = Integer.parseInt(tempItems[0]);
							if(cellTitle.equals(defaultOptions.get(k)))
							{
								Cell cell = row.getCell(1);
								cell.setCellType(Cell.CELL_TYPE_STRING);
								if(cellTitle.equals("store telephone number"))
								{
									String deviceId = calculateDeviceId(cell);
									values.add(deviceId);
								}
								if(cell.getStringCellValue().equals(""))
								{
									String noData = "No Data Provided";
									values.add(noData);
								}
								else
								{
									String value = cell.getStringCellValue();
									values.add(value);
								}
							}
						}
					}*/
				}
			}
			return fnolData;
			
		}
		catch(Exception e)
		{
			return null;
		}
	}
	
	private String calculateDeviceId(Cell regCell, XSSFWorkbook deviceReport)
	{
		try
		{
			if(regCell == null || regCell.getCellType() == Cell.CELL_TYPE_BLANK)
			{
				String noData = "No VRN Provided";
				return noData;
			}
			String carReg = regCell.getStringCellValue();
			String carRegFixed = carReg.replaceAll(" ", "");
			//File file = listFiles[0];
			//FileInputStream fis = new FileInputStream(file);
			//XSSFWorkbook deviceReport = new XSSFWorkbook(fis);
			XSSFSheet sheet = deviceReport.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			
			while(rowIterator.hasNext())
			{
				Row row = rowIterator.next();
				Cell cell = row.getCell(3);
				if(!(cell == null))
				{
					String regBefore = cell.getStringCellValue();
					String regFixed = regBefore.replaceAll(" ", "");
					if(carRegFixed.equals(regFixed))
					{
						Cell cell2 = row.getCell(0);
						String deviceIdString = cell2.getStringCellValue();
						return deviceIdString;
					}
				}
			}
		}
		catch(Exception e)
		{
		}
		
		
		/*File[] listFiles = deviceReportFolder.listFiles();
		
		if(listFiles.length == 0 || listFiles == null)
		{
			return "";
		}
		else
		{
			try
			{
				if(regCell == null || regCell.getCellType() == Cell.CELL_TYPE_BLANK)
				{
					String noData = "No VRN Provided";
					return noData;
				}
				String carReg = regCell.getStringCellValue();
				String carRegFixed = carReg.replaceAll(" ", "");
				//File file = listFiles[0];
				//FileInputStream fis = new FileInputStream(file);
				//XSSFWorkbook deviceReport = new XSSFWorkbook(fis);
				XSSFSheet sheet = deviceReport.getSheetAt(0);
				Iterator<Row> rowIterator = sheet.iterator();
				
				while(rowIterator.hasNext())
				{
					Row row = rowIterator.next();
					Cell cell = row.getCell(3);
					if(!(cell == null))
					{
						String regBefore = cell.getStringCellValue();
						String regFixed = regBefore.replaceAll(" ", "");
						if(carRegFixed.equals(regFixed))
						{
							Cell cell2 = row.getCell(0);
							String deviceIdString = cell2.getStringCellValue();
							return deviceIdString;
						}
					}
				}
			}
			catch(Exception e)
			{
			}
		}*/
		
		return "";
	}
	
	private File createResultsFileWithFnolDataCustom(List<String> resultTitles, List<List<String>> fnolData, List<List<String>> errorData) throws IOException
	{
		String createFileFolderName = createName();
		
		//String path = getCurrentPath();
		
		String path = System.getProperty("user.home")+"/ServerPlayaround/FileUploadDemo/downloadFolder";
		
		
		
		File resultsDownload = new File(path);
		
		System.out.println(resultsDownload.exists());
		
		String resultsDownloadsName = resultsDownload.getAbsolutePath();
		
		String dirLocation = resultsDownloadsName+"\\"+createFileFolderName;
		File newDirectory = new File(dirLocation);
		newDirectory.mkdir();
		
		String fileLocation = dirLocation+"/"+createFileFolderName+".xlsx";
		//String fileLocation = "C:/Users/Matt/ServerPlayaround/FileUploadDemo/resultingFileLocation/test.xlsx";
		
		File resultsWorkbook = new File(fileLocation);
		
		XSSFWorkbook spreadsheet = new XSSFWorkbook();
		XSSFSheet sheet = spreadsheet.createSheet("Results");
		XSSFSheet problemFnols = spreadsheet.createSheet("Problem FNOLs");
		
		Row problemFnolsTitle = problemFnols.createRow(0);
		Cell fnol = problemFnolsTitle.createCell(0);
		Cell reason = problemFnolsTitle.createCell(1);
		fnol.setCellValue("FNOL Title");
		reason.setCellValue("Reason For Fail");
		
		Row titleRow = sheet.createRow(0);
		
		for(String option : resultTitles)
		{
			int colNum = resultTitles.indexOf(option);
			Cell cell = titleRow.createCell(colNum);
			cell.setCellValue(option);
			sheet.autoSizeColumn(0);
		}
		
		for(List<String> unprocessedFnolData : errorData)
		{
			int lastRow = problemFnols.getLastRowNum();
			Row problemRow = problemFnols.createRow(lastRow+1);
			String fnolName = unprocessedFnolData.get(0);
			String fnolReason = unprocessedFnolData.get(1);
			Cell cell1 = problemRow.createCell(0);
			Cell cell2 = problemRow.createCell(1);
			cell1.setCellValue(fnolName);
			cell2.setCellValue(fnolReason);
			problemFnols.autoSizeColumn(0);
			problemFnols.autoSizeColumn(1);
		}
		
		for(List<String> results : fnolData)
		{
			int i = 0;
			int lastRow = sheet.getLastRowNum();
			Row sheetRow = sheet.createRow(lastRow+1);
			for(String result : results)
			{
				Cell cell = sheetRow.createCell(i);
				cell.setCellValue(result);
				sheet.autoSizeColumn(i);
				i++;
			}
		}
		
		try
		{
			FileOutputStream outputStream = new FileOutputStream(resultsWorkbook);
			spreadsheet.write(outputStream);
			outputStream.flush();
			outputStream.close();
		}
		catch(Exception e)
		{
			
		}
		

		return resultsWorkbook;
	}
	
	private String createName()
	{
		DateFormat dateFormat = new SimpleDateFormat("ddMMyyyy HH:mm:ss");
		Date date = new Date();
		String dateString = dateFormat.format(date).toString();
		String dateStringTrim = dateString.replaceAll("\\s", "");
		String dateFinal = dateStringTrim.substring(0, 2)+"-"+dateStringTrim.substring(2, 4)+
				"-"+dateStringTrim.substring(4, 8)+"_"+dateStringTrim.substring(8, dateStringTrim.length());
		String logName = "FNOL_Run_"+dateFinal;
		String logNameNoColons = logName.replaceAll(":", "-");
		return logNameNoColons;
	}
	
	private void reformatCells(File file)
	{
		try
		{
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			
			CellStyle dateStyle = workbook.createCellStyle();
			CreationHelper createHelper = workbook.getCreationHelper();
			dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
			
			CellStyle timeStyle = workbook.createCellStyle();
			timeStyle.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm:ss"));
			
			XSSFSheet sheet = workbook.getSheetAt(0);
			int totalRows = sheet.getLastRowNum();
			
			for(int i = 1; i < totalRows+1; i++)
			{
				Row row = sheet.getRow(i);
				Cell dateCell = row.getCell(3);
				String dateString = dateCell.getStringCellValue();
				Cell timeCell = row.getCell(4);
				String timeString = timeCell.getStringCellValue();
				FormulaEvaluator fev = workbook.getCreationHelper().createFormulaEvaluator();
				dateCell.setCellFormula("VALUE("+dateString+")");
				timeCell.setCellFormula("VALUE("+timeString+")");
				fev.evaluate(dateCell);
				fev.evaluate(timeCell);
				dateCell.setCellStyle(dateStyle);
				timeCell.setCellStyle(timeStyle);
			}
			
			FileOutputStream outputStream = new FileOutputStream(file);
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();
			
			
			
			
			
		}
		catch(Exception e)
		{
			
		}

		
		
		
		
		
		
		
		
		
		
	}
	
	private void downloadFile(HttpServletResponse response, File file) throws ServletException, IOException
	{
		String filePath = file.getAbsolutePath();
		System.out.println(filePath);
		System.out.println(file.exists());
		String fileName = file.getName();
		
		
		response.reset();
		
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setHeader("Content-Disposition",
	                     "attachment;filename="+fileName);
		//ServletContext ctx = getServletContext();
		//InputStream is = ctx.getResourceAsStream(filePath);
		InputStream is = new FileInputStream(file);
			
		int read=0;
		byte[] bytes = new byte[BYTES_DOWNLOAD];
		OutputStream os = response.getOutputStream();
			
		while((read = is.read(bytes))!= -1){
			os.write(bytes, 0, read);
		}
		os.flush();
		os.close();
		is.close();
	}
	

	
}
