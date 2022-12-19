package excelUtils;

import java.io.FileInputStream;
import java.io.IOException;
//import java.util.Scanner;
import java.util.UUID;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Uuid extends ExcelImplications {
	static Logger logs;
	Uuid(){
		logs = LogManager.getLogger(Uuid.class);
	}
	@Override
	void ExcelInit() throws IOException {
		// TODO Auto-generated method stub
		logs.info("Writing Unique IDs into file....");
		input = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(input);
		sheet = workbook.getSheetAt(0);
		int[] cellNum = {3,4,5,6,7,8,9,10,11,195,196};
		for (int i = 1; i <= UserDataManager.users; i++) {
			XSSFRow row = sheet.getRow(i);
			if(row == null) {
				row = sheet.createRow(i);
			}
			for(int j = 0; j<11; j++) {
				XSSFCell cell = row.getCell(cellNum[j]);
				if(cell == null) {
					cell = row.createCell(cellNum[j]);
				}
				UUID uuid = UUID.randomUUID();
				String a = uuid.toString();
				cell.setCellValue(a);
			}
		}
		logs.info("Unique IDs written successfully");
	}
}
