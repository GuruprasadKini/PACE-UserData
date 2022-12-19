package excelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;
import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProductDataManager {
	static FileInputStream fileIn;
	static XSSFWorkbook workbook;
	static FileOutputStream fileOut;
	static FileOutputStream fileOutput;
	static MultiValuedMap<String, String> productInfo;
	public static String[] keys;
	static Logger log;
	public static void writeProductCondition() throws IOException {
		log = LogManager.getLogger(ProductDataManager.class);
		log.info("Writing Product Condition.....");
		FileInputStream fileIn = new FileInputStream("./File/UserData.xlsx");
		Workbook workbook = WorkbookFactory.create(fileIn);
		Sheet sheet1 = workbook.getSheetAt(0);
		productInfo = new ArrayListValuedHashMap<String, String>();
		keys = new String[UserDataManager.users+1];
		int a = (int) Math.round((UserDataManager.users * 0.1)/3);
		int b = (int) Math.round((UserDataManager.users * 0.8)/3);
		for(int i =1; i <= UserDataManager.users; i++) {
			Row row = sheet1.getRow(i);
			if(row == null){
				row = sheet1.createRow(i);
			}
			for(int j =13; j<=14;j++) {
				switch(j) {
				case(13):{
					Cell cell = row.getCell(j);
					DataFormatter formatter1 = new DataFormatter();
					keys[i] = formatter1.formatCellValue(cell).toString();
					break;
				}
				case(14):{
					Cell cell = row.getCell(j);
					if(cell == null){
						cell = row.createCell(i);
					}
					while(i <=a) {
						cell.setCellValue("5");
						productInfo.put(keys[i], i+"-"+"5");
						break;
					}
					while(a < i & i <= 2*(a)) {
						cell.setCellValue("8");
						productInfo.put(keys[i], i+"-"+"8");
						break;
					}
					while(2*(a) < i & i <= 3*(a)) {
						cell.setCellValue("10");
						productInfo.put(keys[i], i+"-"+"10");
						break;
					}
					while(3*(a) < i & i <= b + 3*(a)) {
						cell.setCellValue("25");
						productInfo.put(keys[i], i+"-"+"25");
						break;
					}
					while(3*(a) + b < i & i <= 2*(b) + 3*(a)) {
						cell.setCellValue("40");
						productInfo.put(keys[i], i+"-"+"40");
						break;
					}
					while(3*(a) + 2*(b) < i & i <= 3*(b) + 3*(a)) {
						cell.setCellValue("100");
						productInfo.put(keys[i], i+"-"+"100");
						break;
					}
					while(3*(a) + 3*(b) < i & i <= 3*(b) + 4*(a)) {
						cell.setCellValue("110");
						productInfo.put(keys[i], i+"-"+"110");
						break;
					}
					while(4*(a) + 3*(b) < i & i <= 3*(b) + 5*(a)) {
						cell.setCellValue("150");
						productInfo.put(keys[i], i+"-"+"150");
						break;
					}
					while(5*(a) + 3*(b) < i & i <= 3*(b) + 6*(a)) {
						cell.setCellValue("180");
						productInfo.put(keys[i], i+"-"+"180");
						break;
					}
				}
				}
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
	}
	public static MultiValuedMap<String, String> productIds;
	public static void getProducts(String filePath) throws IOException {
		log.info("Getting Product IDs.....");
		fileIn = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(fileIn);
		UserDataManager.getData(filePath);
		String[] bits;
		String lastOne;
		ArrayList<Integer> limit = new ArrayList<Integer>();
		productIds = new ArrayListValuedHashMap<String, String>();
		for(int i =0; i < UserDataManager.key.length ;i++) {
			XSSFSheet sheet = workbook.getSheet(UserDataManager.key[i]);
			Iterator<String> iterator = productInfo.get(UserDataManager.key[i]).iterator();
			while(iterator.hasNext()) {
				bits = iterator.next().split("-");
				lastOne = bits[bits.length-1];
				limit.add(Integer.parseInt(lastOne));
			}
			int max = Collections.max(limit);
			for(int j = 2; j <= max + 2 ; j++ ) {
				XSSFRow row = sheet.getRow(j);
					XSSFCell cell = row.getCell(1);
					DataFormatter formatter1 = new DataFormatter();
					String productId = formatter1.formatCellValue(cell).toString();
					if(productId=="") {
						cell = row.getCell(2);
						formatter1 = new DataFormatter();
						productId = formatter1.formatCellValue(cell).toString();
					}
					productIds.put(UserDataManager.key[i], productId);
			}
		}
		workbook.close();
	}
	public static void writeProductIds(String filePath) throws EncryptedDocumentException, IOException {
		log.info("Writing Product IDs.....");
		fileIn = new FileInputStream("./File/UserData.xlsx");
		Workbook workbook = WorkbookFactory.create(fileIn);
		Sheet sheet1 = workbook.getSheetAt(0);
		UserDataManager.getData(filePath);
		String[] bits;
		int productCount;
		int rowNums;
		for(int i =0; i < UserDataManager.key.length ;i++) {
			Iterator<String> iterator1 = productInfo.get(UserDataManager.key[i]).iterator();
			while(iterator1.hasNext()) {
				bits = iterator1.next().split("-");
				productCount = Integer.parseInt(bits[1]);
				rowNums = Integer.parseInt(bits[0]);
				Row row = sheet1.getRow(rowNums);
				List<Object> firstNElementsList = productIds.get(UserDataManager.key[i]).stream().limit(productCount).collect(Collectors.toList());
				for(int cellNum =15; cellNum < 15 + productCount; cellNum++) {
					Cell cell = row.getCell(cellNum);
					if(cell == null) {
						cell = row.createCell(cellNum);
					}
					String value = (String)firstNElementsList.get(cellNum-15);
					cell.setCellValue(value);
				}
			}
		}
		fileOut = new FileOutputStream("./File/UserData.xlsx");
		fileOutput = new FileOutputStream("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.xlsx");
		workbook.write(fileOutput);
		fileOutput.close();
		workbook.write(fileOut);
		workbook.close();
		fileIn.close();
		log.info("Product IDs written successfully");
	}	
}
