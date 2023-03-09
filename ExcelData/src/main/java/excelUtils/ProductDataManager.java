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
		int lastIndex = index;
		int perBottlerOrders = userDataManager.users/3;
//		int countKeeper = perBottlerOrders*count;
		DataFormatter formatter1 = new DataFormatter();
		for (int i = index; i < perBottlerOrders + lastIndex; i++) {
			index++;
			row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}

			cell = row.getCell(12);
			String customerId = formatter1.formatCellValue(cell).toString();

			cell = row.getCell(13);
			if (cell == null) {
				cell = row.createCell(13);
			}

			int productCondition = 0;
			if (i <= lastIndex + (int) Math.round((perBottlerOrders * 0.4) / 3)) {
				productCondition = 5;
			} else if (lastIndex + (int) Math.round((perBottlerOrders * 0.4) / 3) < i & i <= lastIndex +  2 * ((int) Math.round((perBottlerOrders * 0.4) / 3))) {
				productCondition = 8;
			} else if(lastIndex + 2 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) < i & i <= lastIndex + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3))){
				productCondition = 10;
			}
			else if(lastIndex + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) < i & i <=lastIndex +  (int) Math.round((perBottlerOrders * 0.5) / 3) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3))) {
				productCondition = 40;
			}
			else if(lastIndex + (int) Math.round((perBottlerOrders * 0.5) / 3) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) < i & i <= lastIndex + 2 * ((int) Math.round((perBottlerOrders * 0.5) / 3)) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3))) {
				productCondition = 70;
			}	
			else if(lastIndex + 2 * ((int) Math.round((perBottlerOrders * 0.5) / 3)) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) < i & i <= lastIndex + 3 * ((int) Math.round((perBottlerOrders * 0.5) / 3)) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3))) {
				productCondition = 100;
			}	
			else if(3 * ((int) Math.round((perBottlerOrders * 0.5) / 3)) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) < i & i <= lastIndex + 3 * ((int) Math.round((perBottlerOrders * 0.5) / 3)) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) + ((int) Math.round((perBottlerOrders * 0.1) / 2))) {
				productCondition = 130;
			}
			else if(lastIndex + 3 * ((int) Math.round((perBottlerOrders * 0.5) / 3)) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) + ((int) Math.round((perBottlerOrders * 0.1) / 2)) < i & i <=lastIndex +  3 * ((int) Math.round((perBottlerOrders * 0.8) / 3)) + 3 * ((int) Math.round((perBottlerOrders * 0.4) / 3)) + 2 * ((int) Math.round((perBottlerOrders * 0.1) / 2))) {
				productCondition = 160;
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
            int randomNo = (int) ((Math.random() * (10 - 0)) + 0);
            for (int j = 2 + randomNo; j <= 2 + randomNo + max; j++) {
            	row = sheet.getRow(j);
            	if(row == null) {
            		row = sheet.getRow(2);
            	}
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
		            cell = row.createCell(j + 14);
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