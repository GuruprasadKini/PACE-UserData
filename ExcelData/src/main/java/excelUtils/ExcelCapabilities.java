package excelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.opencsv.CSVWriter;

public class ExcelCapabilities{
    XSSFWorkbook workbook;
    FileInputStream fileIn;
    FileOutputStream fileOut;
    FileOutputStream fileOutput;
    public void ExcelCreate() {
    	workbook = new XSSFWorkbook();
    }
    
    public void ExcelInit(String filePath) throws IOException{
        fileIn = new FileInputStream(filePath);
        IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
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
     
    public void excelToCSV() throws IOException {
    	// Read the Excel file
    	FileInputStream inputStream = new FileInputStream("./File/UserData.xlsx");
    	workbook = new XSSFWorkbook(inputStream);
    	XSSFSheet sheet = workbook.getSheetAt(0);

    	// Create the CSV file
    	FileWriter fileWriter = new FileWriter("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.csv");
    	CSVWriter csvWriter = new CSVWriter(fileWriter);

    	// Write the data from the Excel file to the CSV file
    	for (Row row : sheet) {
    		String[] data = new String[row.getLastCellNum()];
    		for (int i = 0; i < row.getLastCellNum(); i++) {
    			Cell cell = row.getCell(i);
    			if(cell == null) {
    				break;
    			}
    			data[i] = cell.toString();
    		}
    		csvWriter.writeNext(data);
    	}
    	// Close the CSV file
    	csvWriter.close();
    }
    
}