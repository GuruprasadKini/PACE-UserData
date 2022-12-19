package excelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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
        fileIn = new FileInputStream(filePath);
        workbook = new XSSFWorkbook(fileIn);
    }
    
    public void Destructor() throws IOException{
        fileOut = new FileOutputStream("./File/UserData.xlsx");
        fileOutput = new FileOutputStream("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.xlsx");
        workbook.write(fileOut);
        workbook.write(fileOutput);
        fileOutput.close();
        fileOut.close();
        workbook.close();
    }
    
    public void inputDestructor() throws IOException {
    	workbook.close();
    	fileIn.close();
    }
}
