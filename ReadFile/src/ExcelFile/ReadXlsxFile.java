package ExcelFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXlsxFile {
	
	public void readExcel(String filePath,String fileName,String sheetName) throws IOException{

	    //Create an object of File class to open xlsx file

	    File file =    new File("C:\\ExcelData\\TestData.xlsx");

	    //Create an object of FileInputStream class to read excel file

	    FileInputStream inputStream = new FileInputStream(file);

	    Workbook wb = null;

	    //Find the file extension by splitting file name in substring  and getting only extension name

	    String fileExtensionName = fileName.substring(fileName.indexOf("."));
       
	    //If it is xlsx file then create object of XSSFWorkbook class

	    wb = new XSSFWorkbook(inputStream);



	    //Read sheet inside the workbook by its name

	    Sheet guru99Sheet = wb.getSheet("Feuil1");

	    //Find number of rows in excel file

	    int rowCount = guru99Sheet.getLastRowNum()-guru99Sheet.getFirstRowNum();

	    //Create a loop over all the rows of excel file to read it

	    for (int i = 0; i < rowCount+1; i++) {

	        Row row = guru99Sheet.getRow(i);

	        //Create a loop to print cell values in a row

	        for (int j = 0; j < row.getLastCellNum(); j++) {

	            //Print Excel data in console

	            System.out.print(row.getCell(j).getStringCellValue()+"|| ");

	        }

	        System.out.println("");
	    } 

	    }  

	    //Main function is calling readExcel function to read data from excel file

	    public static void main(String[] args) throws Exception{

	    //Create an object of ReadXlsxFile class

	    ReadXlsxFile objExcelFile = new ReadXlsxFile();

	    //Prepare the path of excel file

	    String filePath = System.getProperty("C:\\ExcelData\\TestData.xlsx")+"\\src\\ExcelFile";

	    //Call read file method of the class to read data

	    objExcelFile.readExcel(filePath,"TestData.xlsx","C:\\ExcelData\\TestData.xlsx");
	    
	    
	    

	    }

	}

	

