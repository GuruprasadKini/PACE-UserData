package excelUtils;

import java.io.IOException;
import java.util.UUID;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Uuid extends ExcelCapabilities {
	UserDataManager userDataManager;
	static Logger logs;
	Uuid(UserDataManager u){
		this.userDataManager = u;
		logs = LogManager.getLogger(Uuid.class);
	}
	void WriteUuid(String filePath) throws IOException{
		logs.info("Writing Unique IDs into file....");
		ExcelInit(filePath);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int[] cellNum = {3,4,5,6,7,8,9,10,11,195,196};
		for (int i = 1; i <= userDataManager.users; i++) {
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
		fileIn.close();
		Destructor();
	}
}
