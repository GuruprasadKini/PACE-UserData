package excelUtils;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UserDataManager {
	static FileInputStream fileIn;
	static XSSFWorkbook workbook;
	static FileOutputStream fileOut;
	static FileOutputStream fileOutput;
	public Scanner threads;
	static Logger logs;
	UserDataManager(){
		threads = new Scanner(System.in);
		System.out.print("Enter Number of users for PACE Performance Test: ");
		users = (threads.nextInt())*2;
		logs = LogManager.getLogger(UserDataManager.class);
	}
	public static int users;
	public static Map<String, String[]> data;
	public static String[] values;
	public static String[] key;
	public static void createFile() throws IOException {
		logs.info("Creating new Excel File......");
		workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("User Data");
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
		fileOut = new FileOutputStream("./File/UserData.xlsx"); 
		fileOutput = new FileOutputStream("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.xlsx");
		workbook.write(fileOut);
		workbook.write(fileOutput);
		fileOut.close();
		fileOutput.close();
		workbook.close();
		logs.info("Excel File has been created");
	}
	public static Map<String, String[]> getData(String filePath) throws IOException {
		FileInputStream inBottler = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(inBottler);
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
		workbook.close();
		inBottler.close();
		return data;
	}
	public static void WriteUserData() throws IOException {
		logs.info("Writing bottler data into the file.....");
		fileIn = new FileInputStream("./File/UserData.xlsx");
		workbook = new XSSFWorkbook(fileIn);
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
		fileOut = new FileOutputStream("./File/UserData.xlsx");
		fileOutput = new FileOutputStream("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.xlsx");
		workbook.write(fileOut);
		workbook.write(fileOutput);
		fileOutput.close();
		fileOut.close();
		workbook.close();
		fileIn.close();
		logs.info("Bottler Data written successfully");
	}
}
//Global variables workbook, sheet, fileinput, fileout