package excelUtils;

import java.awt.HeadlessException;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.bouncycastle.util.Arrays;

public class UserDataManager extends ExcelCapabilities {
	public int users;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	UserDataManager(int users) {
		this.users = users;
//		logs = LogManager.getLogger(UserDataManager.class);
	}

	UserDataManager(UserDataManager u) {
		// Copy Constructor
		this.users = u.users;
	}

	Map<String, String[]> data;
	String[] values;
	public static String[] key;

	public void createFile() throws IOException {
		ExcelCreate();
		sheet = workbook.createSheet("UserData");
		FileReader fileIns = new FileReader("./File/headers.txt");
		BufferedReader read = new BufferedReader(fileIns);
		row = sheet.createRow(0);
		int cellNum = 0;
		String header;
		while ((header = read.readLine()) != null) {
		    cell = row.createCell(cellNum++);
		    cell.setCellValue(header);
		}
		read.close();
		fileIns.close();
		Destructor();
	}

	public void getBottlerData(String filePath) throws IOException {
        	ExcelInit(filePath);
            sheet = workbook.getSheetAt(0);
            data = new HashMap<String, String[]>();
            key = new String[sheet.getLastRowNum()];
            DataFormatter formatter = new DataFormatter();
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                cell = row.getCell(0);
                if (cell == null) {
                    continue;
                }
                key[rowNum - 1] = formatter.formatCellValue(cell).toString();
                values = new String[8];
                for (int cellNum = 0; cellNum < 8; cellNum++) {
                    cell = row.getCell(cellNum);
                    if (cell == null) {
                        continue;
                    }
                    values[cellNum] = formatter.formatCellValue(cell).toString();
                }
                data.put(key[rowNum - 1], values);
            }
            inputDestructor();
        }
    


    public void WriteUserData(String UserCredentialsFile) throws IOException, HeadlessException, UnsupportedFlavorException, InterruptedException {
        GetAuthentication getAuth = new GetAuthentication();
        getAuth.getMobToken(UserCredentialsFile);
//    	getAuth.getWebToken(UserCredentialsFile);
        ExcelInit("./File/UserData.xlsx");
        sheet = workbook.getSheetAt(0);
        int lastIndex = sheet.getLastRowNum() + 1;
        int index = 0;
        int pullNumber = 1;
        int sequenceNumber = 1;
        int perBottlerUsers = users/3;
        Map<String, String> botlerIdToUserId = new HashMap<>();
        botlerIdToUserId.put("5000", "a11196bc-8191-435e-9428-85838d5cea08");
        botlerIdToUserId.put("4200", "272cacfd-7d33-4c4b-9ff9-be046d2432ee");
        botlerIdToUserId.put("4100", "8cc858b2-5429-49f4-981c-2f1aa5f88304");
            for (int rowNum = lastIndex; rowNum < lastIndex + perBottlerUsers; rowNum++) {
            	row = sheet.createRow(rowNum);
                
//            	//Writing WebAuthToken
//    			cell = row.createCell(0);
//    			cell.setCellValue(getAuth.WebToken);

    			//Writing MobAuthToken
    			cell = row.createCell(0);
    			cell.setCellValue(getAuth.MobToken);
                
    			//Writing Bottler ID
                cell = row.createCell(1);
                cell.setCellValue(data.get(key[index])[1]);
                
                //Writing UUIDs
                for (int j = 0; j < 11; j++) {
                    cell = row.createCell(CellNum.values()[j].getValue());
                    cell.setCellValue(getUUID());
                }
                
                //Writing Route ID
                cell = row.createCell(11);
                cell.setCellValue(data.get(key[index])[2]);
                
                //Writing Customer ID
                cell = row.createCell(12);
                cell.setCellValue(key[index]);
                
                //Writing Sequence Number
                cell = row.createCell(196);
                cell.setCellValue(sequenceNumber++);
                
                //Writing Pull Number
                cell = row.createCell(197);
                cell.setCellValue(pullNumber);
                
                //Writing Form number
                cell = row.createCell(198);
                cell.setCellValue("PF" + rowNum);
                
                //Writing User ID
                cell = row.createCell(199);
                cell.setCellValue(botlerIdToUserId.get(data.get(key[index])[1]));
                
                //Writing 
                cell = row.createCell(200);
                cell.setCellValue(data.get(key[index])[3]);
                
                cell = row.createCell(201);
    			cell.setCellValue(data.get(key[index])[4]);
    
    			cell = row.createCell(202);
    			cell.setCellValue(data.get(key[index])[5]);
    
    			cell = row.createCell(203);
    			cell.setCellValue(data.get(key[index])[6]);
    
    			cell = row.createCell(204);
    			cell.setCellValue(data.get(key[index])[7]);
    			index++;
    			if (index >= key.length) {
    				index = 0;
    				pullNumber++;
    			}
            }
            Arrays.fill(key, null);
            data.clear();
            Destructor();
            fileIn.close();
            }
    
    private static String getUUID() {
    	return UUID.randomUUID().toString();
    }

    enum CellNum {
    	CELL_2(2), CELL_3(3), CELL_4(4), CELL_5(5), CELL_6(6), CELL_7(7), CELL_8(8), CELL_9(9), CELL_10(10), CELL_194(194), CELL_195(195);

    	private int value;

    	CellNum(int value) {
    		this.value = value;
    	}

    	public int getValue() {
    		return value;
    	}
    }

}