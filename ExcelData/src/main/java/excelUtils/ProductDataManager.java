package excelUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ProductDataManager extends ExcelCapabilities {
	static MultiValuedMap<String, String> productInfo;
	public static String[] keys;
	public int userNum;
	UserDataManager userDataManager = new UserDataManager(userNum);
	static Logger log;
	int index = 1;
	
	ProductDataManager(UserDataManager u) {
		this.userDataManager = u;
	}

	public void writeProductCondition() throws IOException {

		log = LogManager.getLogger(ProductDataManager.class);
		log.info("Writing Product Condition.....");
		ExcelInit("./File/UserData.xlsx");
		XSSFSheet sheet1 = workbook.getSheetAt(0);
		productInfo = new ArrayListValuedHashMap<String, String>();
		int perBottlerUsers = userDataManager.users / 3;
		int a = (int) Math.round((userDataManager.users * 0.1) / 3);
		int b = (int) Math.round((userDataManager.users * 0.8) / 3);
		int c = 2 * (a);
		int d = 3 * (a);
		int e = b + 3 * (a);
		int f = 2 * (b) + 3 * (a);
		int g = 3 * (b) + 3 * (a);
		int h = 3 * (b) + 4 * (a);
		int p = 3 * (b) + 5 * (a);
		int q = 3 * (b) + 6 * (a);
		int lastIndex = index;
		DataFormatter formatter1 = new DataFormatter();
		keys = new String[perBottlerUsers + lastIndex + 1];
		for (int i = index; i < perBottlerUsers + lastIndex; i++) {
			index++;
			Row row = sheet1.getRow(i);
			if (row == null) {
				row = sheet1.createRow(i);
			}

			Cell customerIdCell = row.getCell(13);
			String customerId = formatter1.formatCellValue(customerIdCell).toString();

			Cell productConditionCell = row.getCell(14);
			if (productConditionCell == null) {
				productConditionCell = row.createCell(14);
			}

			int productCondition = 0;
			if (i <= a) {
				productCondition = 5;
			} else if (a < i & i <= c) {
				productCondition = 8;
			} else if(c < i & i <= d){
				productCondition = 10;
			}
			else if(d < i & i <= e) {
				productCondition = 25;
			}
			else if(e < i & i <= f) {
				productCondition = 40;
			}	
			else if(f < i & i <= g) {
				productCondition = 100;
			}	
			else if(g < i & i <= h) {
				productCondition = 110;
			}
			else if(h < i & i <= p) {
				productCondition = 150;
			}
			else if(p < i & i <= q) {
				productCondition = 180;
			}
			productConditionCell.setCellValue(productCondition);
			productInfo.put(customerId, i + "-" + productCondition);
		}
		Destructor();
		fileIn.close();
	}

	public MultiValuedMap<String, String> productIds;

	public void getProducts(String filePath) throws IOException {
		log.info("Getting Product IDs.....");
		ExcelInit(filePath);
		userDataManager.getBottlerData(filePath);
		String[] bits;
		String lastOne;
		ArrayList<Integer> limit = new ArrayList<Integer>();
		productIds = new ArrayListValuedHashMap<String, String>();
		for (int i = 0; i < UserDataManager.key.length; i++) {
			XSSFSheet sheet = workbook.getSheet(UserDataManager.key[i]);
			Iterator<String> iterator = productInfo.get(UserDataManager.key[i]).iterator();
			while (iterator.hasNext()) {
				bits = iterator.next().split("-");
				lastOne = bits[bits.length - 1];
				limit.add(Integer.parseInt(lastOne));
			}
			int max = Collections.max(limit);
			for (int j = 2; j <= 2 + max; j++) {
				XSSFRow row = sheet.getRow(j);
				XSSFCell cell = row.getCell(1);
				DataFormatter formatter1 = new DataFormatter();
				String productId = formatter1.formatCellValue(cell).toString();
				if (Objects.isNull(productId)) {
					cell = row.getCell(2);
					formatter1 = new DataFormatter();
					productId = formatter1.formatCellValue(cell).toString();
				}
				productIds.put(UserDataManager.key[i], productId);	
			}
		}
		inputDestructor();
	}

	public void writeProductIds(String filePath) throws EncryptedDocumentException, IOException {
		log.info("Writing Product IDs.....");
		ExcelInit("./File/UserData.xlsx");
		XSSFSheet sheet = workbook.getSheetAt(0);
		userDataManager.getBottlerData(filePath);

		for (int i = 0; i < UserDataManager.key.length; i++) {
		    Iterator<String> iterator = productInfo.get(UserDataManager.key[i]).iterator();
		    while (iterator.hasNext()) {
		        String[] bits = iterator.next().split("-");
		        int productCount = Integer.parseInt(bits[1]);
		        int rowNum = Integer.parseInt(bits[0]);
		        XSSFRow row = sheet.getRow(rowNum);
		        if(row == null) {
		            row = sheet.createRow(rowNum);
		        }
		        List<String> firstNElementsList = ((List<String>) productIds.get(UserDataManager.key[i])).subList(0, productCount);
		        for (int j = 0; j < productCount; j++) {
		            XSSFCell cell = row.createCell(j + 15);
		            cell.setCellValue(firstNElementsList.get(j));
		        }
		    }
		}
		Destructor();
		fileIn.close();
		productIds.clear();
		log.info("Product IDs written successfully");
	}
}