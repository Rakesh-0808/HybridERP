package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.commons.compress.harmony.pack200.NewAttribute;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtile
{
	XSSFWorkbook wb;
	public ExcelFileUtile(String Excelpath) throws Throwable
	{
		FileInputStream fi = new FileInputStream(Excelpath);
		wb = new XSSFWorkbook(fi);  
	}
	public int rowcount(String Sheetname)
	{
		return wb.getSheet(Sheetname).getLastRowNum();
	}
	public String cellread(String sheetname, int row, int col) {
		String data="";
		if(wb.getSheet(sheetname).getRow(row).getCell(col).getCellType()==CellType.NUMERIC)
		{
			int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(col).getNumericCellValue();
			data =String.valueOf(celldata);
		}
		else 
		{
			data = wb.getSheet(sheetname).getRow(row).getCell(col).getStringCellValue();
		}
		return data;  
	}
	public void setcelldata(String sheetname,int row,int col,String status,String writeExcel) throws Throwable {
		XSSFSheet WS= wb.getSheet(sheetname);
		XSSFRow RC = WS.getRow(row);
		XSSFCell CL = RC.createCell(col);
		CL.setCellValue(status);
		if(status.equalsIgnoreCase("pass"))
		{
			XSSFCellStyle cell = wb.createCellStyle();
			XSSFFont Fo = wb.createFont();
			Fo.setColor(IndexedColors.GREEN.getIndex());
			Fo.setBold(true);
			cell.setFont(Fo);
			RC.getCell(col).setCellStyle(cell);
		}
		else if(status.equalsIgnoreCase("fail"))
		{
			XSSFCellStyle cell = wb.createCellStyle();
			XSSFFont Fo = wb.createFont();
			Fo.setColor(IndexedColors.RED.getIndex());
			Fo.setBold(true);
			cell.setFont(Fo);
			RC.getCell(col).setCellStyle(cell);
		}
		else if(status.equalsIgnoreCase("Blocked"))
		{
			XSSFCellStyle cell = wb.createCellStyle();
			XSSFFont Fo = wb.createFont();
			Fo.setColor(IndexedColors.BLUE.getIndex());
			Fo.setBold(true);
			cell.setFont(Fo);
			RC.getCell(col).setCellStyle(cell);
		}
        FileOutputStream fo = new FileOutputStream(writeExcel);
        wb.write(fo);
	}

	
} 
