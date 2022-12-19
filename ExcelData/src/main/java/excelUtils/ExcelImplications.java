package excelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public abstract class ExcelImplications {
	String filePath = "./File/UserData.xlsx";
	FileInputStream input;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	public void ExcelStart() throws IOException {
		input = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(input);
	}
	abstract void ExcelInit() throws IOException;
	public void ExcelClose() throws IOException {
		FileOutputStream f = new FileOutputStream(filePath);
		FileOutputStream fileOut = new FileOutputStream("./File/UserData.xlsx");
		workbook.write(fileOut);
		workbook.write(f);
		fileOut.close();
		f.close();
		workbook.close();
		input.close();
	}
}
class ExcelUtils {
	String filePath = "C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.xlsx";
	FileInputStream input;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFRow row;
	public String getCellData(int rowNum, int cellNum) throws IOException {
		row = sheet.getRow(rowNum);
		XSSFCell cell = row.getCell(cellNum);
		String data;
		DataFormatter formatter = new DataFormatter();
		try {
			data = formatter.formatCellValue(cell);
		}
		catch(Exception e){
			data = "";
		}
		workbook.close();
		input.close();
		return data;
	}
}