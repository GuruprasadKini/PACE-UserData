package excelUtils;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
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
    	   FileInputStream fileIn = new FileInputStream("./File/UserData.xlsx");
           @SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
           XSSFSheet sheet = workbook.getSheetAt(0);
           fileIn.close();

           // Write the CSV file
           BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(
               new FileOutputStream("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.csv"), StandardCharsets.UTF_8));
           writer.write('\ufeff'); // add BOM for Excel compatibility

           for (Row row : sheet) {
               for (int i = 0; i < row.getLastCellNum(); i++) {
                   Cell cell = row.getCell(i);
                   if (cell == null) {
                       writer.write("");
                   } else if (cell.getCellType() == CellType.NUMERIC) {
                       writer.write(String.valueOf(cell.getNumericCellValue()));
                   } else if (cell.getCellType() == CellType.STRING) {
                       writer.write(cell.getStringCellValue());
                   }
                   writer.write(",");
               }
               writer.newLine();
           }

           writer.flush();
           writer.close();
       }
}