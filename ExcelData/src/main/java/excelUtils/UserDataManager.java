package excelUtils;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.HashMap;
import java.util.Map;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class UserDataManager extends ExcelCapabilities {
	public int users;
	static Logger logs;
	UserDataManager(int users){
		this.users = users;
		logs = LogManager.getLogger(UserDataManager.class);
	}
	UserDataManager(UserDataManager u){
		//Copy Constructor
		this.users = u.users;
	}
	public static Map<String, String[]> data;
	public static String[] values;
	public static String[] key;
	
	public void createFile() throws IOException {
		logs.info("Creating new Excel File......");
		ExcelCreate();
		XSSFSheet sheet = workbook.createSheet("UserData");
		//read from txt and make headers dynamic 
		FileReader fileIn = new FileReader("./File/headers.txt");
		BufferedReader read = new BufferedReader(fileIn);
		int newLine = 0;
		String[] header = new String[210];
		while((header[newLine] = read.readLine())!=null) {
			newLine++;
		}
		read.close();
		for(int rowNum =0; rowNum < 1; rowNum++) {
			XSSFRow row = sheet.createRow(rowNum);
			for(int cellNum = 0; cellNum < Array.getLength(header); cellNum++ ) {
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
		for(int rowNum =1; rowNum< sheet1.getLastRowNum()+1; rowNum++) {
			XSSFRow row = sheet1.getRow(rowNum);
			XSSFCell cell = row.getCell(0);
			DataFormatter formatter1 = new DataFormatter();
			key[rowNum-1] = formatter1.formatCellValue(cell).toString();
			values= new String[8];
			for(int cellNum = 0; cellNum < 8; cellNum++) {
				cell = row.getCell(cellNum);
				DataFormatter formatter2 = new DataFormatter();
				values[cellNum] = formatter2.formatCellValue(cell).toString();
			}
			data.put(key[rowNum-1], values);
		}
		inputDestructor();
	}
	public void WriteUserData() throws IOException {
		logs.info("Writing bottler data into the file.....");
		ExcelInit("./File/UserData.xlsx");
		XSSFSheet sheet1 = workbook.getSheetAt(0);
		//get the customerID's and print in excel column and use that column as key and get rest of the data 
		//Print customerId's multiple times 
		//Make it dynamic for all bottlerIds
		int index = 0;
		int lastIndex = sheet1.getLastRowNum()+1;
		int pullNumber = 1;
		int sequenceNumber = 1;
		String[] userId = {"a11196bc-8191-435e-9428-85838d5cea08","272cacfd-7d33-4c4b-9ff9-be046d2432ee","8cc858b2-5429-49f4-981c-2f1aa5f88304"};
		int perBottlerUsers = Math.round(users/3);
		for(int rowNum = lastIndex; rowNum < lastIndex + perBottlerUsers; rowNum++) {
			XSSFRow row = sheet1.createRow(rowNum);
			for(int cellNum =0; cellNum <= 205; cellNum++ ) {
				XSSFCell cell = row.createCell(cellNum);
				switch(cellNum) {
				case 2:{
					cell.setCellValue(data.get(key[index])[1]);
					break;
				}
				case 12:{
					cell.setCellValue(data.get(key[index])[2]);
					break;
				}
				case 13:{
					cell.setCellValue(key[index]);
					break;
				}
				//Module 9:
				case 197:{
					cell.setCellValue(sequenceNumber++);
					break;
				}
				//Module 10:
				case 198:{
					cell.setCellValue(pullNumber);
					break;
				}
				//Module 11:
				case 199:{
					cell.setCellValue("PF"+rowNum);
					break;
				}
				//Module 8:
				case 200:{
					if(data.get(key[index])[1].equals("5000")) {
						cell.setCellValue(userId[0]);
					}
					else if(data.get(key[index])[1].equals("4200")) {
						cell.setCellValue(userId[1]);
					}
					else {
						cell.setCellValue(userId[2]);
					}
					break;
				}
				case 201:{
					cell.setCellValue(data.get(key[index])[3]);
					break;
				}
				case 202:{
					cell.setCellValue(data.get(key[index])[4]);
					break;
				}
				case 203:{
					cell.setCellValue(data.get(key[index])[5]);
					break;
				}
				case 204:{
					cell.setCellValue(data.get(key[index])[6]);
					break;
				}
				case 205:{
					cell.setCellValue(data.get(key[index])[7]);
					break;
				}
				}
			}
			index++;
			if(index >= key.length) {
				index = 0;
				pullNumber++;
			}			
		}
		fileIn.close();
		Destructor();
		logs.info("Bottler Data written successfully");
	}
}