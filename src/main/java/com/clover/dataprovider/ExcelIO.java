package com.clover.dataprovider;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.testng.annotations.DataProvider;

public class ExcelIO {
		
	String resultFilePath =null;
	String sheetName;
	
	@DataProvider(name = "TestData") 
	public static Object[][] Desktop() throws Exception {	
		Object[][] retObjArr = ReadAllDatawithDataProvider("TestData/CloverTestData.xls","QA");	
		return retObjArr;
	}
	
	public void AddColumn(String FileName, String sheetname, ArrayList <String> Data)throws Exception
	{	
		HSSFWorkbook WorkBook = new HSSFWorkbook();		
		if(!new File(FileName).exists())
		{
			FileOutputStream fileOut = new FileOutputStream(FileName);//To check if file exits if not it will create Excel file			
			HSSFSheet Sheet = WorkBook.createSheet(sheetname);			
			WorkBook.write(fileOut);
			int k = Sheet.getPhysicalNumberOfRows();					
			//int i = 0;
			for(String AddData : Data)//Add array of data into one row
			{	
				HSSFRow row1 = Sheet.createRow((short)k);
				HSSFCell cellA1 = row1.createCell(0);				
				cellA1.setCellType(HSSFCell.CELL_TYPE_STRING);			
				cellA1.setCellValue((AddData));				
				k++;			
			}
			fileOut.close();
		}
		else //In existing Excel file add data into a new sheet
		{
			FileInputStream filein = new FileInputStream(new File(FileName));

			WorkBook = new HSSFWorkbook(filein);			
			int NS = WorkBook.getNumberOfSheets() - 1;
			HSSFSheet Sheet = null;
			for(int q=0; q<=NS; q++)
			{				
				if(WorkBook.getSheetName(q).equals(sheetname))
				{
					Sheet = WorkBook.getSheet(sheetname);
					break;
				}
				else if(q == NS)
				{
					Sheet = WorkBook.createSheet((sheetname));						
				}					
			}		
			int k = Sheet.getPhysicalNumberOfRows();						
			
			
			for(String AddData : Data)//Add array of data into one row
			{	
				HSSFRow row1 = Sheet.createRow((short)k);
				HSSFCell cellA1 = row1.createCell(0);				
				cellA1.setCellType(HSSFCell.CELL_TYPE_STRING);			
				cellA1.setCellValue((AddData));				
				k++;			
			}
			filein.close();
		}		
		FileOutputStream outFile =new FileOutputStream(new File(FileName));

	    WorkBook.write(outFile);
		outFile.close();
		FileName = null;
		sheetname = null;		
	}
	
	public List<String> AddtoExcel(String resultFile, String sheetname, ArrayList <String> Data) throws Exception //To create and add data to the excel sheet
	{
		int i = 0,k=0;
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		if (!new File(resultFilePath).exists())
		{
		  new File(resultFilePath).mkdirs();//To check if directory exists if not it will create directory
		}		  
		if(!(new File(resultFile).exists()))
		{		
			File file = new File(resultFilePath, String.valueOf(resultFile));
			file.createNewFile();
			FileOutputStream fileOut = new FileOutputStream(resultFile);//To check if file exits if not it will create Excel file	
			
			HSSFSheet Sheet = WorkBook.createSheet(sheetname);
			WorkBook.write(fileOut);			
			HSSFRow row1 = Sheet.createRow((short)k);
			for(String AddData : Data)//Add array of data into one row
			{			
				HSSFCell cellA1 = row1.createCell(i);
				cellA1.setCellType(HSSFCell.CELL_TYPE_STRING);			
				cellA1.setCellValue((AddData));			
				i++;			
			}
			fileOut.close();			
		}
		else //In existing Excel file add data into a new sheet
		{
			FileInputStream filein = new FileInputStream(new File(resultFilePath));		
			HSSFSheet Sheet = null;
			WorkBook = new HSSFWorkbook(filein);
			int NS = WorkBook.getNumberOfSheets() - 1;
			for(int q=0; q<=NS; q++)
			{				
				if(WorkBook.getSheetName(q).equals(sheetname))
				{
					Sheet = WorkBook.getSheet(sheetname);
					k = Sheet.getPhysicalNumberOfRows();
					break;
				}
				else if(q == NS)
				{
					Sheet = WorkBook.createSheet((sheetname));						
				}				
			}		
			i=0;
			HSSFRow row1 = Sheet.createRow((short)k);
			for(String AddData : Data)//Add array of data into one row
			{			
				HSSFCell cellA1 = row1.createCell(i);
				cellA1.setCellType(HSSFCell.CELL_TYPE_STRING);			
				cellA1.setCellValue((AddData));			
				i++;			
			}
			filein.close();
		}		
		k++;
		FileOutputStream outFile =new FileOutputStream(new File(resultFile));
	    WorkBook.write(outFile);
		outFile.close();
		return null;
	}
	
	
	public ArrayList<String> AddRow(String FileName, String sheetname, ArrayList <String> Data)throws Exception {	
		
	HSSFWorkbook WorkBook = new HSSFWorkbook();
	File file = new File(FileName);
	if(!(file.exists())) {
			FileOutputStream fileOut = new FileOutputStream(file);	//To check if file exits if not it will create Excel file			
			HSSFSheet Sheet = WorkBook.createSheet(sheetname);
			//CellStyle style1 = WorkBook.createCellStyle();
			//CellStyle style2 = WorkBook.createCellStyle();
			//Font font = WorkBook.createFont();
			WorkBook.write(fileOut);
			int k = Sheet.getPhysicalNumberOfRows();						
			HSSFRow row1 = Sheet.createRow((short)k);			
			int i = 0;
			for(String AddData : Data)//Add array of data into one row
			{	
				if(AddData != null)
				{
					HSSFCell cellA1 = row1.createCell(i);
					/*if(AddData.equalsIgnoreCase("Pass"))
	    			{
	    				font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		    			font.setColor(IndexedColors.GREEN.getIndex());
		    			style1.setFont(font);
		 				cellA1.setCellStyle(style1);
	    			}
	    			else if(AddData.equalsIgnoreCase("Fail"))
	    			{
	    				style2.setFillForegroundColor(IndexedColors.RED.getIndex());
	    				style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    				style2.setFont(font);
	    				cellA1.setCellStyle(style2);
	    			}*/
					cellA1.setCellType(HSSFCell.CELL_TYPE_STRING);			
					cellA1.setCellValue((AddData));					
					i++;
				}
			}
			fileOut.close();
		}
		else //In existing Excel file add data into a new sheet
			
		{
			FileInputStream filein = new FileInputStream(file); 
			WorkBook = new HSSFWorkbook(filein);
			//CellStyle style1 = WorkBook.createCellStyle();
			//CellStyle style2 = WorkBook.createCellStyle();
			//Font font = WorkBook.createFont();
			//fileOut = new FileOutputStream(TestDataFile);
			int NS = WorkBook.getNumberOfSheets() - 1;
			HSSFSheet Sheet = null;
			for(int q=0; q<=NS; q++)
			{				
				if(WorkBook.getSheetName(q).equals(sheetname))
				{
					Sheet = WorkBook.getSheet(sheetname);
					break;
				}
				else if(q == NS)
				{
					Sheet = WorkBook.createSheet((sheetname));						
				}					
			}		
			int k = Sheet.getPhysicalNumberOfRows();						
			HSSFRow row1 = Sheet.createRow((short)k);
			int i = 0;
			for(String AddData : Data)//Add array of data into one row
			{	
				if(AddData != null)
				{
					HSSFCell cellA1 = row1.createCell(i);
					/*if(AddData.equalsIgnoreCase("Pass"))
	    			{
	    				font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		    			font.setColor(IndexedColors.GREEN.getIndex());
		    			style1.setFont(font);
		 				cellA1.setCellStyle(style1);
	    			}
	    			else if(AddData.equalsIgnoreCase("Fail"))
	    			{    				
	    				style2.setFillForegroundColor(IndexedColors.RED.getIndex());
	    				style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    				style2.setFont(font);
	    				cellA1.setCellStyle(style2);
	    			}*/
					cellA1.setCellType(HSSFCell.CELL_TYPE_STRING);			
					cellA1.setCellValue((AddData));
					
					i++;	
				}
			}
			filein.close();
		}		
		FileOutputStream outFile =new FileOutputStream(file); 
	    WorkBook.write(outFile);
		outFile.close();
		FileName = null;
		sheetname = null;		
		return null;		
	}
	
	
	
	public void RenameLastSheet(String File, String NewName)throws Exception
	{
		FileInputStream fileOut = null;
		HSSFWorkbook workbook;
		
		fileOut = new FileInputStream(new File(File));
		workbook = new HSSFWorkbook(fileOut);			
		int NS = workbook.getNumberOfSheets() - 1;		
		for(int e=0; e<=NS;e++)
		{
			if(workbook.getSheetName(e).equalsIgnoreCase(NewName))
			{
				workbook.setSheetName(workbook.getNumberOfSheets()-1, NewName+e);
				break;
			}
			else if(e==workbook.getNumberOfSheets()-1)
			{
				workbook.setSheetName(workbook.getNumberOfSheets()-1, NewName);
			}
		}
		fileOut.close();
		FileOutputStream outFile =new FileOutputStream(new File(File));
		workbook.write(outFile);
		outFile.close();
	}
	
	public boolean SearchSheet(String File, String SheetName, String SearchKeyword) throws Exception
	{
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		FileInputStream file = new FileInputStream(new File(File));
		WorkBook = new HSSFWorkbook(file);		
		HSSFSheet Sheet = WorkBook.getSheet(SheetName);	
		Iterator <Row> rowIterator = Sheet.iterator();
		while(rowIterator.hasNext()) //Iterate until the last row
		{
			Row row = rowIterator.next();	       
	        Iterator <Cell> cellIterator = row.cellIterator();	        
	    	while(cellIterator.hasNext()) //Iterate until the last cell
	    	{
	    		Cell cell = cellIterator.next();
	    		int strg = cell.getCellType();
	    		if(strg ==1) //if the cell type is String
	    		{
	    			if(cell.getRichStringCellValue().getString().equals(SearchKeyword))
		    		{
	    				return true;
		    		}	    			
	    		}
	    		else if(strg ==0)//If the cell type is numeric
	    		{
	    			
	    			double nume = cell.getNumericCellValue();
	    			
	    			NumberFormat formatter = new DecimalFormat("#0");
	    			String cellval = formatter.format(nume);	    			
	    			if(cellval.equals(SearchKeyword))
	    			{
	    				return true;
	    			}	    			
	    		}
	    	}
		}
		return false;
	}
	
	
	public ArrayList<String> ReadAllData(String Filename, String Sheetname)throws Exception
	{
		
		ArrayList<String> ReadData = new ArrayList <String>();
		
		FileInputStream file = new FileInputStream(resultFilePath);	     
	    //Get the workbook instance for XLS file 
	    HSSFWorkbook workbook = new HSSFWorkbook(file);  //Get first sheet from the workbook
	    HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);	    
	    HSSFSheet sheet = workbook.getSheet(Sheetname);	     
	    //Iterate through each rows from first sheet
	    Iterator<Row> rowIterator = sheet.iterator();	    
	    while(rowIterator.hasNext()) 
	    {
	       	Row row = rowIterator.next();	        
	        //For each row, iterate through each columns
	        Iterator<Cell> cellIterator = row.cellIterator();
	        while(cellIterator.hasNext()) 
	        {
	           Cell cell = cellIterator.next();	          
	           switch(cell.getCellType()) 
	           {
	                case Cell.CELL_TYPE_BOOLEAN: 
	                	Boolean b = cell.getBooleanCellValue();
	                	ReadData.add(b.toString());
	                	break;
	                case Cell.CELL_TYPE_NUMERIC:	                	
	                	double nume = cell.getNumericCellValue();
	                	DecimalFormat DoubleFormat = new DecimalFormat("#.##");
	                	String ss = DoubleFormat.format(nume);
		    			ReadData.add(ss);	                    
	                    break;
	                case Cell.CELL_TYPE_STRING:	                	
	                	ReadData.add(cell.getStringCellValue());	                    
	                    break;
	                case Cell.CELL_TYPE_FORMULA:
	                	if(cell.getCachedFormulaResultType()==1)
	                	{
	                		ReadData.add(cell.getStringCellValue());
	                	}
	            }
	        }	        
	    }	   
	    file.close();
	    FileOutputStream out = new FileOutputStream(new File(Filename));
	    workbook.write(out);
	    out.close();
	    return ReadData;
	}
		
	public static String[][] ReadAllDatawithDataProvider(String Filename, String Sheetname) throws Exception 
	{
		
		String[][] tabArray = null;
		FileInputStream file = new FileInputStream(new File(Filename));
		//Get the workbook instance for XLS file 
		HSSFWorkbook WorkBook1 = new HSSFWorkbook(file);
		//Get first sheet from the workbook
		HSSFSheet Sheet1 = WorkBook1.getSheet(Sheetname);
		//Iterate through each rows from first sheet
		Iterator <Row> rowIterator = Sheet1.iterator();
		Row row = rowIterator.next();
		row = rowIterator.next();
		int numberofColumns = row.getLastCellNum();
		int numberofRows = Sheet1.getPhysicalNumberOfRows();
	  	tabArray=new String[numberofRows-1][numberofColumns];
		int ci=0;				
		for(int i=1;i<numberofRows;i++,ci++)
		{
			int cj=0;
		  
			for(int j=0;j<numberofColumns;j++,cj++)
			{
				int strg = row.getCell(j).getCellType();
	    		if(strg ==1) //if the cell type is String
	    		{
	    			tabArray[ci][cj] = row.getCell(j).getStringCellValue();		    			    			
	    		}
	    		else if(strg ==0)//If the cell type is numeric
	    		{
	    			double nume = row.getCell(j).getNumericCellValue();	    			
	    			NumberFormat formatter = new DecimalFormat("#0");
	    			tabArray[ci][cj] = formatter.format(nume);	    			
	    		}				
			}
			if(i==numberofRows-1)
			{
				break;
			}
			row = rowIterator.next();	
		}
		file.close();
		FileOutputStream out = new FileOutputStream(new File(Filename));
		WorkBook1.write(out);
		out.close();
		return tabArray;
	}
	
	public ArrayList<String> GetSpecificRow(String File, String SheetName, String SearchData ) throws Exception //To get specific row by providing any value in that row
	{
		ArrayList<String> Getrow = new ArrayList <String>();
		FileInputStream file = new FileInputStream(new File(File));
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		WorkBook = new HSSFWorkbook(file);		
		HSSFSheet Sheet = WorkBook.getSheet(SheetName);	
		Iterator <Row> rowIterator = Sheet.iterator();
		String ID = null;
		while(rowIterator.hasNext()) //Iterate until the last row
		{
			Row row = rowIterator.next();	       
	        Iterator <Cell> cellIterator = row.cellIterator();	        
	    	while(cellIterator.hasNext()) //Iterate until the last cell
	    	{
	    		Cell cell = cellIterator.next();
	    		int strg = cell.getCellType();
	    		if(strg ==1) //if the cell type is String
	    		{
	    			if(cell.getRichStringCellValue().getString().equals(SearchData))
		    		{		    			
		    			int i= row.getPhysicalNumberOfCells();
		    			for(int k=0; k<i; k++)
		    			{		    				
		    					Cell celldata = row.getCell(k);		    					
		    					ID =  celldata.getRichStringCellValue().getString();                   			
		    					Getrow.add(ID);		    				
		                }		    			
		    		}
	    		}
	    		
	    		else if(strg ==0)//If the cell type is numeric
	    		{
	    			
	    			double nume = cell.getNumericCellValue();
	    			
	    			NumberFormat formatter = new DecimalFormat("#0");
	    			String cellval = formatter.format(nume);	    			
	    			if(cellval.equals(SearchData))
	    			{
	    				int i= row.getPhysicalNumberOfCells();
	    				for(int k=0; k<i; k++)
	    				{
	    					Cell celldata = row.getCell(k);
	    					double numeval = celldata.getNumericCellValue();
	    					String numeval2 = new Double(numeval).toString();		    				                  			
		                    Getrow.add(numeval2);	                   
	    				}	    			
	    			}
	    		}	    		
	    	}	    	
	    }
		if(!(Getrow.isEmpty()))
		{			
			return Getrow;
		}
		return null;		
	}
	
	public void UpateSpecificRow(String TestDataFile, String sheetname, String ParamName, String ParamValue) throws Exception
	{	
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		FileInputStream file = new FileInputStream(new File(TestDataFile));
		//Get the workbook instance for XLS file 
		WorkBook = new HSSFWorkbook(file);
		//CellStyle style1 = WorkBook.createCellStyle();
		//Font font = WorkBook.createFont();
		HSSFSheet Sheet = WorkBook.getSheet(sheetname);
		
		Iterator <Row> rowIterator = Sheet.iterator();
		
		while(rowIterator.hasNext()) //Iterate until the last row
		{
			Row row = rowIterator.next();	     
	        Iterator <Cell> cellIterator = row.cellIterator();	        
	    	while(cellIterator.hasNext()) 
	    	{
	    		Cell cell = cellIterator.next(); //Iterate until the last cell
	    		int strg = cell.getCellType();
	    		if(strg == 1)//if the cell type is string
	    		{	    			
		    		if(cell.getStringCellValue().equalsIgnoreCase(ParamName))
		    		{
		    			int k = row.getLastCellNum();		    			
		    			cell = row.createCell((short) k);		    		   
		    			/*if(ParamValue.equalsIgnoreCase("Pass"))
		    			{
		    				 font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			    			 font.setColor(IndexedColors.GREEN.getIndex());
		    			}
		    			else if(ParamValue.equalsIgnoreCase("Fail"))
		    			{
		    				style1.setFillForegroundColor(IndexedColors.RED.getIndex());
		    				style1.setFillPattern(CellStyle.SOLID_FOREGROUND);		    				
		    			}*/
		    			cell.setCellValue(ParamValue);
		    		   // style1.setFont(font);
		    		  //  cell.setCellStyle(style1);
		    			file.close();
		    		    FileOutputStream outFile =new FileOutputStream(new File(TestDataFile));
		    		    WorkBook.write(outFile);
		    			outFile.close();
		    			return;
		    		}
	    		}
	    		else if(strg ==0) //if the cell type is numeric
	    		{
	    			double nume = cell.getNumericCellValue();
	    			
	    			NumberFormat formatter = new DecimalFormat("#0");
	    			String cellval = formatter.format(nume);	    			
	    			if(cellval.equals(ParamName))
	    			{
	    				int k = row.getLastCellNum();
	    				row.createCell(k).setCellValue(ParamValue);
	    				file.close();
	    			    FileOutputStream outFile =new FileOutputStream(new File(TestDataFile));
	    			    WorkBook.write(outFile);
	    				outFile.close();
	    				return;	    				    					    				
	    			}
	    		}
	    	}
		}		
	}
	
	public void UpdateException(String ResultFile, String Exception) throws Exception //If the code fails from exception create/update excel with the exception
	{
		
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		HSSFSheet Sheet = null;
		/*if (!new File(resultFilePath).exists())
		{
		  new File(resultFilePath).mkdirs();//To check if directory exists if not it will create directory
		}*/
		
		if((!new File(resultFilePath).exists()))	//To check if file exists if not file will be created
		{
			FileOutputStream fileOut = new FileOutputStream(new File(resultFilePath));
			Sheet = WorkBook.createSheet("Exception");
			WorkBook.write(fileOut);			
		}
		else
		{			
			FileInputStream filein = new FileInputStream(new File(resultFilePath));//To create sheet called Exception
			Sheet = null;
			WorkBook = new HSSFWorkbook(filein);
			int NS = WorkBook.getNumberOfSheets() - 1;
			for(int q=0; q<=NS; q++)
			{				
				if(WorkBook.getSheetName(q).equals("Exception"))
				{
					Sheet = WorkBook.getSheet("Exception");
					break;
				}
				else if(q == NS)
				{
					Sheet = WorkBook.createSheet("Exception");										
				}				
			}				
		}
		Sheet = WorkBook.getSheet("Exception");
		int g = Sheet.getLastRowNum(); //update data in the last row of the sheet
		g = g+1;
		HSSFRow row1 = Sheet.createRow((short)g);		
		HSSFCell cellA1 = row1.createCell(0);
		cellA1.setCellType(HSSFCell.CELL_TYPE_STRING);			
		cellA1.setCellValue((Exception));
		FileOutputStream outFile =new FileOutputStream(new File(ResultFile));
	    WorkBook.write(outFile);
		outFile.close();
	}
	
	public void ClearLastCell(String TestDataFile, String sheetname,String ParamName) throws Exception
	{
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		FileInputStream file = new FileInputStream(new File(TestDataFile));
		//Get the workbook instance for XLS file 
		WorkBook = new HSSFWorkbook(file);
		
		HSSFSheet Sheet = WorkBook.getSheet(sheetname);
		
		Iterator <Row> rowIterator = Sheet.iterator();
		
		while(rowIterator.hasNext()) //Iterate until the last row
		{
			Row row = rowIterator.next();	     
	        Iterator <Cell> cellIterator = row.cellIterator();	        
	    	while(cellIterator.hasNext()) 
	    	{
	    		Cell cell = cellIterator.next(); //Iterate until the last cell
	    		int strg = cell.getCellType();
	    		if(strg == 1)//if the cell type is string
	    		{	    			
		    		if(cell.getStringCellValue().equals(ParamName))
		    		{
		    			int k = row.getLastCellNum();
		    			k = k-1;
		    			if(k>0)
		    			{
		    				row.removeCell(row.getCell(k));		    			
			    			file.close();
			    		    FileOutputStream outFile =new FileOutputStream(new File(TestDataFile));
			    		    WorkBook.write(outFile);
			    			outFile.close();
			    			return;
		    			}		    			
		    		}
	    		}
	    		else if(strg ==0) //if the cell type is numeric
	    		{
	    			double nume = cell.getNumericCellValue();
	    			
	    			NumberFormat formatter = new DecimalFormat("#0");
	    			String cellval = formatter.format(nume);	    			
	    			if(cellval.equals(ParamName))
	    			{
	    				int k = row.getLastCellNum();
		    			k = k-1;
		    			if(k>0)
		    			{
		    				row.removeCell(row.getCell(k));		    			
			    			file.close();
			    			FileOutputStream outFile =new FileOutputStream(new File(TestDataFile));
			    			WorkBook.write(outFile);
			    			outFile.close();
			    			return;
		    			}
	    			}
	    		}
	    	}
		}	
	}
	
	public void ClearRow(String Filename, String Sheetname, String Param)throws Exception
	{		
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		FileInputStream file = new FileInputStream(new File(Filename));
		//Get the workbook instance for XLS file 
		WorkBook = new HSSFWorkbook(file);
		
		HSSFSheet Sheet = WorkBook.getSheet(Sheetname);
		
		Iterator <Row> rowIterator = Sheet.iterator();
		
		while(rowIterator.hasNext()) //Iterate until the last row
		{
			Row row = rowIterator.next();	     
	        Iterator <Cell> cellIterator = row.cellIterator();	        
	    	while(cellIterator.hasNext()) 
	    	{
	    		Cell cell = cellIterator.next(); //Iterate until the last cell
	    		int strg = cell.getCellType();
	    		if(strg == 1)//if the cell type is string
	    		{    			
		    		if(cell.getStringCellValue().equals(Param))
		    		{
		    			int k= row.getPhysicalNumberOfCells();
		    			for(int q=1; q<k; q++)
		    			{
		    				row.removeCell(row.getCell(q));		    			
			    		}
		    			break;
		    		}
	    		}
	    		else if(strg ==0) //if the cell type is numeric
	    		{
	    			double nume = cell.getNumericCellValue();
	    			
	    			NumberFormat formatter = new DecimalFormat("#0");
	    			String cellval = formatter.format(nume);	    			
	    			if(cellval.equals(Param))
	    			{
	    				int k= row.getPhysicalNumberOfCells();
	    				for(int q=1; q<k; q++)
		    			{
		    				row.removeCell(row.getCell(q));				    			
			    		}
	    				break;
	    			}
	    		}
	    	}							
		}
		file.close();
		FileOutputStream outFile =new FileOutputStream(new File(Filename));
		WorkBook.write(outFile);
		outFile.close();
	}
	
	public int GetNumberofRows(String Filename, String Sheetname) throws Exception
	{
		FileInputStream file = new FileInputStream(new File(Filename));
		//Get the workbook instance for XLS file 
		HSSFWorkbook WorkBook1 = new HSSFWorkbook(file);
		//Get first sheet from the workbook
		HSSFSheet Sheet1 = WorkBook1.getSheet(Sheetname);
		//Iterate through each rows from first sheet
		//Iterator <Row> rowIterator = Sheet1.iterator();
		//Row row = rowIterator.next();
	//	row = rowIterator.next();		
		int numberofRows = Sheet1.getPhysicalNumberOfRows();
		return numberofRows;
	}
	
	public void FinalFormat(String Filename)throws Exception
	{		 
		HSSFWorkbook WorkBook = new HSSFWorkbook();	
		FileInputStream filein = new FileInputStream(new File(Filename));		
		WorkBook = new HSSFWorkbook(filein);
		HSSFSheet Sheet1;
		
		int NS = WorkBook.getNumberOfSheets() - 1;
		for(int q=0; q<=NS; q++)
		{		
			String SName = WorkBook.getSheetName(q);
			Sheet1 = WorkBook.getSheet(SName);			
			Iterator <Row> rowIterator = Sheet1.iterator();
			while(rowIterator.hasNext()) //Iterate until the last row
			{
				Row row = rowIterator.next();	       
		        Iterator <Cell> cellIterator = row.cellIterator();	        
		    	while(cellIterator.hasNext()) //Iterate until the last cell
		    	{
		    		Cell cell = cellIterator.next();
		    		if(cell.getRichStringCellValue().getString().equals("Pass"))
		    		{	  
		    			CellStyle style2 = WorkBook.createCellStyle();
		    			Font font2 = WorkBook.createFont();
		    			font2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		    			font2.setColor(IndexedColors.GREEN.getIndex());
		    			style2.setFont(font2);
		    			cell.setCellStyle(style2);
		    		}
	    			else if(cell.getRichStringCellValue().getString().equals("Fail"))
	    			{
	    				CellStyle style1 = WorkBook.createCellStyle();
	    				Font font = WorkBook.createFont();
	    				font.setColor(HSSFColor.BLACK.index);
	    				style1.setFillForegroundColor(IndexedColors.RED.getIndex());
	    				style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    				cell.setCellStyle(style1);
	    			}
		    				    		
		    	}
			}
		}
		FileOutputStream outFile =new FileOutputStream(new File(Filename));
	    WorkBook.write(outFile);
		outFile.close();
	}
	
	public void UpdateCellNumber(String Filename, String Sheetname, String Param, int Row, int Col)throws Exception
	{		
		HSSFWorkbook WorkBook = new HSSFWorkbook();
		FileInputStream file = new FileInputStream(new File(Filename));
		//Get the workbook instance for XLS file 
		WorkBook = new HSSFWorkbook(file);
		//CellStyle style1 = WorkBook.createCellStyle();
		//Font font = WorkBook.createFont();
		HSSFSheet Sheet = WorkBook.getSheet(Sheetname);
		Row row = Sheet.getRow(Row);
		Cell cell = row.createCell(Col-1);
		cell.setCellValue(Param);	
		FileOutputStream outFile =new FileOutputStream(new File(Filename));
	    WorkBook.write(outFile);
		outFile.close();
	}
	
	public void SetSheetOrder(String Filename, String SheetName, int Position)throws Exception
	{
		HSSFWorkbook WorkBook = new HSSFWorkbook();	
		FileInputStream filein = new FileInputStream(new File(Filename));		
		WorkBook = new HSSFWorkbook(filein);
		WorkBook.setSheetOrder(SheetName, Position);
		FileOutputStream outFile =new FileOutputStream(new File(Filename));
	    WorkBook.write(outFile);
		outFile.close();
	}
	
	public void DeleteFile(String Filename) throws Exception
	{
		File file = new File(Filename);
		if(file.exists())
		{
			file.delete();
		}		
	}
}
	
	