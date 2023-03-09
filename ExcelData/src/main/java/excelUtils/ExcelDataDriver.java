package excelUtils;

import java.util.Scanner;

public class ExcelDataDriver {
//	static Logger logs;
	public static void main(String[] args) {
		Scanner threads = new Scanner(System.in);
		System.out.print("Enter number of virtual users for PACE Performance Test: ");
		int userNum = (int)(threads.nextInt()*2.5);
		if(userNum<3) {
			userNum = 3;
		}
		try {
			
			UserDataManager users = new UserDataManager(userNum);
			ProductDataManager productData = new ProductDataManager(users);
			users.createFile();
			//User 1
			users.getBottlerData("./File/5000_AMLProductIdList_new.xlsx");
			users.WriteUserData("./File/5000_UserCredentials.xlsx");
			productData.writeProductCondition();
			productData.getProducts("./File/5000_AMLProductIdList_new.xlsx");
			productData.writeProductIds("./File/5000_AMLProductIdList_new.xlsx");
			System.out.println("Data for Bottler - 5000 has been written");
		    //User 2
//			users.getBottlerData("./File/4100_AMLProductIdList_new.xlsx");
//			users.WriteUserData("./File/4100_UserCredentials.xlsx");
//			productData.writeProductCondition();
//			productData.getProducts("./File/4100_AMLProductIdList_new.xlsx");
//			productData.writeProductIds("./File/4100_AMLProductIdList_new.xlsx");
//			System.out.println("Data for Bottler - 4100 has been written");
//          //User 3
//			users.getBottlerData("./File/4200_AMLProductIdList_new.xlsx");
//			users.WriteUserData("./File/4200_UserCredentials.xlsx");
//			productData.writeProductCondition();
//			productData.getProducts("./File/4200_AMLProductIdList_new.xlsx");
//			productData.writeProductIds("./File/4200_AMLProductIdList_new.xlsx");
//			System.out.println("Data for Bottler - 4200 has been written");
			ExcelCapabilities excelCapabilities = new ExcelCapabilities();
			excelCapabilities.excelToCsv();
			threads.close();
		}
		catch(Exception e) {
			// handle any exception that occurs
			e.printStackTrace();
		}
	}
}
