package excelUtils;

import java.awt.HeadlessException;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class UserDataManager extends ExcelCapabilities {
	public int users;
	static Logger logs;

	UserDataManager(int users) {
		this.users = users;
		logs = LogManager.getLogger(UserDataManager.class);
	}

	UserDataManager(UserDataManager u) {
		// Copy Constructor
		this.users = u.users;
	}

	public static Map<String, String[]> data;
	public static String[] values;
	public static String[] key;

	public void createFile() throws IOException {
		logs.info("Creating new Excel File......");
		ExcelCreate();
		XSSFSheet sheet = workbook.createSheet("UserData");
		// read from txt and make headers dynamic
		FileReader fileIn = new FileReader("./File/headers.txt");
		BufferedReader read = new BufferedReader(fileIn);
		int newLine = 0;
		String[] header = new String[210];
		while ((header[newLine] = read.readLine()) != null) {
			newLine++;
		}
		read.close();
		for (int rowNum = 0; rowNum < 1; rowNum++) {
			XSSFRow row = sheet.createRow(rowNum);
			for (int cellNum = 0; cellNum < Array.getLength(header); cellNum++) {
				XSSFCell cell = row.createCell(cellNum);
				cell.setCellValue(header[cellNum]);
			}
		}
		Destructor();
		logs.info("Excel File has been created");
	}

	public void getBottlerData(String filePath) throws IOException {
		ExcelInit(filePath);
		XSSFSheet sheet1 = workbook.getSheetAt(0);
		data = new HashMap<String, String[]>();
		key = new String[sheet1.getLastRowNum()];
		for (int rowNum = 1; rowNum < sheet1.getLastRowNum() + 1; rowNum++) {
			XSSFRow row = sheet1.getRow(rowNum);
			XSSFCell cell = row.getCell(0);
			DataFormatter formatter1 = new DataFormatter();
			key[rowNum - 1] = formatter1.formatCellValue(cell).toString();
			values = new String[8];
			for (int cellNum = 0; cellNum < 8; cellNum++) {
				cell = row.getCell(cellNum);
				DataFormatter formatter2 = new DataFormatter();
				values[cellNum] = formatter2.formatCellValue(cell).toString();
			}
			data.put(key[rowNum - 1], values);
		}
		inputDestructor();
	}
	
	public void WriteUserData(String UserCredentialsFile) throws IOException, HeadlessException, UnsupportedFlavorException, InterruptedException {
		GetAuthentication getAuth = new GetAuthentication();
		getAuth.getWebToken(UserCredentialsFile); 
		getAuth.getMobToken(UserCredentialsFile);
		logs.info("Writing bottler data and UUIDs into the file.....");
		ExcelInit("./File/UserData.xlsx");
		XSSFSheet sheet1 = workbook.getSheetAt(0);
		// get the customerID's and print in excel column and use that column as key and
		// get rest of the data
		// Print customerId's multiple times
		// Make it dynamic for all bottlerIds
		int index = 0;
		int lastIndex = sheet1.getLastRowNum() + 1;
		int pullNumber = 1;
		int sequenceNumber = 1;
		int[] cellNum = {3,4,5,6,7,8,9,10,11,195,196};
		String[] userId = { "a11196bc-8191-435e-9428-85838d5cea08", "272cacfd-7d33-4c4b-9ff9-be046d2432ee",
				"8cc858b2-5429-49f4-981c-2f1aa5f88304" };
		int perBottlerUsers = users / 3;
		for (int rowNum = lastIndex; rowNum < lastIndex + perBottlerUsers; rowNum++) {
			XSSFRow row = sheet1.createRow(rowNum);
			XSSFCell cell;
			//Writing WebAuthToken
			cell = row.createCell(0);
			cell.setCellValue(getAuth.WebToken);

			//Writing MobAuthToken
			cell = row.createCell(1);
			cell.setCellValue(getAuth.MobToken);

			cell = row.createCell(2);
			cell.setCellValue(data.get(key[index])[1]);
			
			//Writing UUID's
			for(int j = 0; j<11; j++) {
				cell = row.getCell(cellNum[j]);
				if(cell == null) {
					cell = row.createCell(cellNum[j]);
				}
				UUID uuid = UUID.randomUUID();
				String a = uuid.toString();
				cell.setCellValue(a);
			}
			
			cell = row.createCell(12);
			cell.setCellValue(data.get(key[index])[2]);

			cell = row.createCell(13);
			cell.setCellValue(key[index]);

			// Module 9:
			cell = row.createCell(197);
			cell.setCellValue(sequenceNumber++);

			// Module 10:
			cell = row.createCell(198);
			cell.setCellValue(pullNumber);

			// Module 11:
			cell = row.createCell(199);
			cell.setCellValue("PF" + rowNum);

			// Module 8:
			cell = row.createCell(200);
			if (data.get(key[index])[1].equals("5000")) {
				cell.setCellValue(userId[0]);
			} else if (data.get(key[index])[1].equals("4200")) {
				cell.setCellValue(userId[1]);
			} else {
				cell.setCellValue(userId[2]);
			}

			cell = row.createCell(201);
			cell.setCellValue(data.get(key[index])[3]);

			cell = row.createCell(202);
			cell.setCellValue(data.get(key[index])[4]);

			cell = row.createCell(203);
			cell.setCellValue(data.get(key[index])[5]);

			cell = row.createCell(204);
			cell.setCellValue(data.get(key[index])[6]);

			cell = row.createCell(205);
			cell.setCellValue(data.get(key[index])[7]);
		index++;
		if (index >= key.length) {
			index = 0;
			pullNumber++;
		}
		}
		data.clear();
		fileIn.close();
		Destructor();
		logs.info("Bottler Data and UUIDs written successfully");
	}
}
