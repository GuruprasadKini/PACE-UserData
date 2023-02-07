package excelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelCapabilities{
    XSSFWorkbook workbook;
    FileInputStream fileIn;
    FileOutputStream fileOut;
    FileOutputStream fileOutput;
    public void ExcelCreate() {
    	workbook = new XSSFWorkbook();
    }
    
    public void ExcelInit(String filePath) throws IOException{
    	IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
        fileIn = new FileInputStream(filePath);
        workbook = new XSSFWorkbook(fileIn);
    }
    
    public void Destructor() throws IOException{
        fileOut = new FileOutputStream("./File/UserData.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }
    
    public void inputDestructor() throws IOException {
    	workbook.close();
    	fileIn.close();
    }
    
    public void excelToCsv() throws IOException {
    	fileIn = new FileInputStream("./File/UserData.xlsx");
        workbook = new XSSFWorkbook(fileIn);
        XSSFSheet sheet = workbook.getSheetAt(0);
        fileIn.close();

        // Write the CSV file
        FileOutputStream fos = new FileOutputStream("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.csv");
        OutputStreamWriter osw = new OutputStreamWriter(fos, StandardCharsets.UTF_8);
        PrintWriter pw = new PrintWriter(osw);
        for (Row row : sheet) {
        	  for (int i = 0; i < row.getLastCellNum(); i++) {
        	    Cell cell = row.getCell(i);
        	    if (cell == null) {
        	      pw.print("");
        	    } else if (cell.getCellType() == CellType.NUMERIC) {
                    pw.print(cell.getNumericCellValue());
                } else if (cell.getCellType() == CellType.STRING) {
                  pw.print(cell.getStringCellValue());
                }
        	    pw.print(",");
        	  }
        	  pw.println();
        	}
            
        pw.flush();
        pw.close();
        osw.close();
        fos.close();
    }
}