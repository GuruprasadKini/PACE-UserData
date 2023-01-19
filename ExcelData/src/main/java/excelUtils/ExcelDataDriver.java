package excelUtils;

import java.util.Scanner;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class ExcelDataDriver {
	static Logger logs;
	public static void main(String[] args) {
		logs = LogManager.getLogger(ExcelDataDriver.class);
		Scanner threads = new Scanner(System.in);
		System.out.print("Enter Number of users for PACE Performance Test: ");
		int userNum = (threads.nextInt())*2;
		try {
			UserDataManager users = new UserDataManager(userNum);	 
			users.createFile();
			users.getBottlerData("./File/662477_5000_AMLProductIdList.xlsx");
			users.WriteUserData("./File/Florida_UserCredentials.xlsx");
			users.getBottlerData("./File/583349_4100_AMLProductIdList.xlsx");
			users.WriteUserData("./File/Canada_UserCredentials.xlsx");
			users.getBottlerData("./File/681328_4200_AMLProductIdList.xlsx");
			users.WriteUserData("./File/Swire_UserCredentials.xlsx");
			//need to include Liberty as well
			ProductDataManager productData = new ProductDataManager(users);
			productData.writeProductCondition();
			//change logic to include 4 bottlers and writing ProductCondition for 4 
			productData.getProducts("./File/662477_5000_AMLProductIdList.xlsx");
			productData.writeProductIds("./File/662477_5000_AMLProductIdList.xlsx");
			productData.getProducts("./File/681328_4200_AMLProductIdList.xlsx");
			productData.writeProductIds("./File/681328_4200_AMLProductIdList.xlsx");
			productData.getProducts("./File/583349_4100_AMLProductIdList.xlsx");
			productData.writeProductIds("./File/583349_4100_AMLProductIdList.xlsx");
			Uuid uuid = new Uuid(users);
			uuid.WriteUuid("./File/UserData.xlsx");
			ExcelCapabilities excelCapabilities = new ExcelCapabilities();
			excelCapabilities.excelToCSV();
			threads.close();
		}
		catch(Exception e) {
			// handle any exception that occurs
			e.printStackTrace();
		}
	}
}
