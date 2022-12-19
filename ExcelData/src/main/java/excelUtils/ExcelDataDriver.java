package excelUtils;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class ExcelDataDriver {
	static Logger logs;
	public static void main(String[] args) {
		logs = LogManager.getLogger(ExcelDataDriver.class);
		try {
			new UserDataManager();
			UserDataManager.createFile();
			UserDataManager.getData("./File/662477_5000_AMLProductIdList.xlsx");
			UserDataManager.WriteUserData();
			UserDataManager.getData("./File/583349_4100_AMLProductIdList.xlsx");
			UserDataManager.WriteUserData();
			UserDataManager.getData("./File/681328_4200_AMLProductIdList.xlsx");
			UserDataManager.WriteUserData();
			ProductDataManager.writeProductCondition();
			ProductDataManager.getProducts("./File/662477_5000_AMLProductIdList.xlsx");
			ProductDataManager.writeProductIds("./File/662477_5000_AMLProductIdList.xlsx");
			ProductDataManager.getProducts("./File/681328_4200_AMLProductIdList.xlsx");
			ProductDataManager.writeProductIds("./File/681328_4200_AMLProductIdList.xlsx");
			ProductDataManager.getProducts("./File/583349_4100_AMLProductIdList.xlsx");
			ProductDataManager.writeProductIds("./File/583349_4100_AMLProductIdList.xlsx");
			Uuid uuid = new Uuid();
			uuid.ExcelStart();
			uuid.ExcelInit();
			uuid.ExcelClose();
		}
		catch(FileNotFoundException e){
			logs.info("The Excel File is open, please close the file");
		}
		catch(IOException e) {
			e.printStackTrace();
		}
		catch(NullPointerException e) {
			e.printStackTrace();
		}
	}
}
