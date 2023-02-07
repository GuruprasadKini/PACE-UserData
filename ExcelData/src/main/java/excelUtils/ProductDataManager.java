package excelUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ProductDataManager extends ExcelCapabilities {
	MultiValuedMap<String, String> productInfo;
	String[] keys;
	int userNum;
	UserDataManager userDataManager = new UserDataManager(userNum);
	int index = 1;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	ProductDataManager(UserDataManager u) {
		this.userDataManager = u;
	}

	public void writeProductCondition() throws IOException {
//		log = LogManager.getLogger(ProductDataManager.class);
//		log.info("Writing Product Condition.....");
		ExcelInit("./File/UserData.xlsx");
		sheet = workbook.getSheetAt(0);
		productInfo = new ArrayListValuedHashMap<String, String>();
		int perBottlerUsers = userDataManager.users/3;
		int lastIndex = index;
		DataFormatter formatter1 = new DataFormatter();
		keys = new String[userDataManager.users + lastIndex + 1];
		for (int i = index; i < perBottlerUsers + lastIndex; i++) {
			index++;
			row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}

			cell = row.getCell(13);
			String customerId = formatter1.formatCellValue(cell).toString();

			cell = row.getCell(14);
			if (cell == null) {
				cell = row.createCell(14);
			}

			int productCondition = 0;
			if (i <= (int) Math.round((userDataManager.users * 0.1) / 3)) {
				productCondition = 5;
			} else if ((int) Math.round((userDataManager.users * 0.1) / 3) < i & i <= 2 * ((int) Math.round((userDataManager.users * 0.1) / 3))) {
				productCondition = 8;
			} else if(2 * ((int) Math.round((userDataManager.users * 0.1) / 3)) < i & i <= 3 * ((int) Math.round((userDataManager.users * 0.1) / 3))){
				productCondition = 10;
			}
			else if(3 * ((int) Math.round((userDataManager.users * 0.1) / 3)) < i & i <= (int) Math.round((userDataManager.users * 0.8) / 3) + 3 * ((int) Math.round((userDataManager.users * 0.1) / 3))) {
				productCondition = 25;
			}
			else if((int) Math.round((userDataManager.users * 0.8) / 3) + 3 * ((int) Math.round((userDataManager.users * 0.1) / 3)) < i & i <= 2 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 3 * ((int) Math.round((userDataManager.users * 0.1) / 3))) {
				productCondition = 40;
			}	
			else if(2 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 3 * ((int) Math.round((userDataManager.users * 0.1) / 3)) < i & i <= 3 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 3 * ((int) Math.round((userDataManager.users * 0.1) / 3))) {
				productCondition = 100;
			}	
			else if(3 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 3 * ((int) Math.round((userDataManager.users * 0.1) / 3)) < i & i <= 3 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 4 * ((int) Math.round((userDataManager.users * 0.1) / 3))) {
				productCondition = 110;
			}
			else if(3 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 4 * ((int) Math.round((userDataManager.users * 0.1) / 3)) < i & i <= 3 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 5 * ((int) Math.round((userDataManager.users * 0.1) / 3))) {
				productCondition = 150;
			}
			else if(3 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 5 * ((int) Math.round((userDataManager.users * 0.1) / 3)) < i & i <= 3 * ((int) Math.round((userDataManager.users * 0.8) / 3)) + 6 * ((int) Math.round((userDataManager.users * 0.1) / 3))) {
				productCondition = 180;
			}
			cell.setCellValue(productCondition);
			productInfo.put(customerId, i + "-" + productCondition);
		}
		Destructor();
		fileIn.close();
	}

	public MultiValuedMap<String, String> productIds;

    public void getProducts(String filePath) throws IOException {
//        log.info("Getting Product IDs.....");
        ExcelInit(filePath);
        userDataManager.getBottlerData(filePath);
        String[] bits;
        String lastOne;
        ArrayList<Integer> limit = new ArrayList<Integer>();
        productIds = new ArrayListValuedHashMap<String, String>();
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < UserDataManager.key.length; i++) {
            sheet = workbook.getSheet(UserDataManager.key[i]);
            Iterator<String> iterator = productInfo.get(UserDataManager.key[i]).iterator();
            while (iterator.hasNext()) {
                bits = iterator.next().split("-");
                lastOne = bits[bits.length - 1];
                limit.add(Integer.parseInt(lastOne));
            }
            int max = Integer.MIN_VALUE;
            for (Integer l : limit) {
                if (l > max) {
                    max = l;
                }
            }
            for (int j = 2; j <= 2 + max; j++) {
            	row = sheet.getRow(j);
            	cell = row.getCell(1);
            	String productId = formatter.formatCellValue(cell).toString();
            	if (Objects.isNull(productId)) {
            		cell = row.getCell(2);
            		formatter = new DataFormatter();
            		productId = formatter.formatCellValue(cell).toString();
            	}
            	productIds.put(UserDataManager.key[i], productId);
            }
        }
        inputDestructor();
    }


	public void writeProductIds(String filePath) throws EncryptedDocumentException, IOException {
//		log.info("Writing Product IDs.....");
		ExcelInit("./File/UserData.xlsx");
		sheet = workbook.getSheetAt(0);
		userDataManager.getBottlerData(filePath);

		for (int i = 0; i < UserDataManager.key.length; i++) {
		    Iterator<String> iterator = productInfo.get(UserDataManager.key[i]).iterator();
		    while (iterator.hasNext()) {
		        String[] bits = iterator.next().split("-");
		        int productCount = Integer.parseInt(bits[1]);
		        int rowNum = Integer.parseInt(bits[0]);
		        row = sheet.getRow(rowNum);
		        if(row == null) {
		            row = sheet.createRow(rowNum);
		        }
		        List<String> firstNElementsList = ((List<String>) productIds.get(UserDataManager.key[i])).subList(0, productCount);
		        for (int j = 0; j < productCount; j++) {
		            cell = row.createCell(j + 15);
		            cell.setCellValue(firstNElementsList.get(j));
		        }
		    }
		}
		Destructor();
		fileIn.close();
		productIds.clear();
		productInfo.clear();
//		log.info("Product IDs written successfully");
	}
}