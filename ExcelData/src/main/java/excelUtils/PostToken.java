package excelUtils;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;

public class PostToken extends ExcelImplications{
	public static String data;
	public void okHttp() throws IOException {
		OkHttpClient client = new OkHttpClient().newBuilder().build();

		MediaType mediaType = MediaType.parse("application/json");

		@SuppressWarnings("deprecation")
		RequestBody body = RequestBody.create(mediaType,
				"{\r\n    \"USERNAME\" : \"admin\",\r\n    \"PASSWORD\" : \"hotwax@786\"\r\n}");

		Request request = new Request.Builder().url("https://dev-apps.hotwax.io/api/login").method("POST", body)

				.addHeader("Authorization", "Basic YWRtaW46aG90d2F4QDc4Ng==")

				.addHeader("Content-Type", "application/json")

				.addHeader("Cookie", "JSESSIONID=BE84742A6E5ECCA6D3E1F171991BF279.jvm1")

				.build();

		Response response = client.newCall(request).execute();
		String arr = response.body().string().toString();
		String array[] = arr.split(",");

		Map<String, String> hashMap = new TreeMap<String, String>();
		for (int i = 0; i < array.length; i++) {
			String p = array[i];
			String o[] = p.split(":");
			hashMap.put(o[0], o[1]);
		}
		String f = hashMap.get("{\"token\"");
		String l[] = f.split("\"");
		data = l[1];
	}
	@Override
	void ExcelInit() throws IOException {
		// TODO Auto-generated method stub
		input = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(input);
		sheet = workbook.getSheetAt(0);
		int rowNum;
		Scanner c = new Scanner(System.in);
		System.out.print("Enter number of threads: ");
		int userNum = c.nextInt();
		c.close();
		for (rowNum = 1; rowNum <= userNum ; rowNum++) {
			XSSFRow row = sheet.getRow(rowNum);
			XSSFCell cell = row.getCell(0);
			cell.setCellValue(data);
		}
	}
	public static void main(String[] args) throws IOException {
		PostToken writeExcel = new PostToken();
		writeExcel.okHttp();
		writeExcel.ExcelStart();
		writeExcel.ExcelInit();
		writeExcel.ExcelClose();
	}
	
}
