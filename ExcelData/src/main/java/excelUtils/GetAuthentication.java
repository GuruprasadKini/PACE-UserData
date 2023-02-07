package excelUtils;

import java.awt.HeadlessException;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
import java.time.Duration;

import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import io.github.bonigarcia.wdm.WebDriverManager;
import okhttp3.MediaType;
import okhttp3.MultipartBody;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;
import okhttp3.ResponseBody;

public class GetAuthentication extends ExcelCapabilities {
	public String MobToken;
	public String WebToken;
	UserDataManager userDataManager;
	static Logger logs;
	WebDriver driver;
	WebDriverWait wait;

	GetAuthentication() {
		// Empty
//		logs = LogManager.getLogger(GetAuthentication.class);
	}

	GetAuthentication(UserDataManager u) {
		this.userDataManager = u;
//		logs = LogManager.getLogger(GetAuthentication.class);
	}
	public void WebDriverActions(String URL, String username, String password) {
		
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");
		driver = new ChromeDriver(options);
		wait = new WebDriverWait(driver, Duration.ofSeconds(30));
		driver.get(URL);
		// Enter Username
		WebElement email = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@type='email']")));
		email.sendKeys(username);
		// click submit button
		WebElement submit = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@type='submit']")));
		submit.click();
		// Enter password
		WebElement passwd = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@name='passwd']")));
		passwd.sendKeys(password);
		// Click sign in button
		WebElement signIn = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@value='Sign in']")));
		signIn.click();
		// Click next button for more information required
		WebElement btn_nxt = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@value='Next']")));
		btn_nxt.click();
		// Skip microsoft Authenticator setup
		WebElement btn_skip = wait
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[text()='Skip setup']")));
		btn_skip.click();
	}
	
	public void getMobToken(String fileName) throws InterruptedException, JsonMappingException, JsonProcessingException, IOException {
		ExcelInit(fileName);
		XSSFSheet sheet = workbook.getSheet("Merch");
		XSSFRow row = sheet.getRow(0);
		//get username
		XSSFCell cell = row.getCell(0);
		DataFormatter formatter2 = new DataFormatter();
		String username = formatter2.formatCellValue(cell).toString();
		//get password
		cell = row.getCell(1);
		String password = formatter2.formatCellValue(cell).toString();
		fileIn.close();
		inputDestructor();
		WebDriverActions("https://login.microsoftonline.com/9120276d-6bfd-47fe-a09f-cdcf486f545f/oauth2/v2.0/authorize?client_id=82db3499-d34d-44ea-89a5-6315c747b1ee&scope=https%3A%2F%2Fgraph.windows.net%2F.default&response_type=code&response_mode=query&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient",
				username, password);
		Thread.sleep(3000);
		String code = driver.getCurrentUrl().split("=")[1].split("&")[0];
		driver.quit();
		OkHttpClient client = new OkHttpClient().newBuilder().build();
		@SuppressWarnings("unused")
		MediaType mediaType = MediaType.parse("text/plain");
		RequestBody body = new MultipartBody.Builder().setType(MultipartBody.FORM)
				.addFormDataPart("grant_type", "authorization_code")
				.addFormDataPart("client_id", "82db3499-d34d-44ea-89a5-6315c747b1ee")
				.addFormDataPart("redirect_uri", "https://login.microsoftonline.com/common/oauth2/nativeclient")
				.addFormDataPart("scope", "https://graph.windows.net/.default offline_access")
				.addFormDataPart("code", code).build();
		Request request = new Request.Builder()
				.url("https://login.microsoftonline.com/9120276d-6bfd-47fe-a09f-cdcf486f545f/oauth2/v2.0/token")
				.method("POST", body)
				.addHeader("Cookie",
						"fpc=Aj6Ez2jabgVBnGLo_-t6I04jayNuAQAAAIEFUtsOAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd")
				.build();
		Response response = client.newCall(request).execute();
		ResponseBody responseBody = response.body();
		ObjectMapper mapper = new ObjectMapper();
		JsonNode rootNode = mapper.readTree(responseBody.string());
		JsonNode specificNode = rootNode.path("access_token");
		MobToken = specificNode.toString().substring(1, specificNode.toString().length() - 1);
		if (MobToken.contains("eY")) {
		System.out.println("Successfully extracted Mobile Authentication Token");
		}
	}
	
	public void getWebToken(String fileName) throws HeadlessException, UnsupportedFlavorException, IOException, InterruptedException {
		ExcelInit(fileName);
		XSSFSheet sheet = workbook.getSheet("Supervisor");
		XSSFRow row = sheet.getRow(0);
		//get username
		XSSFCell cell = row.getCell(0);
		DataFormatter formatter2 = new DataFormatter();
		String username = formatter2.formatCellValue(cell).toString();
		//get password
		cell = row.getCell(1);
		String password = formatter2.formatCellValue(cell).toString();
		fileIn.close();
		inputDestructor();
		WebDriverActions("https://paceq.c1nacloud.com/web/auth", username, password);
		WebElement btn_yes = wait
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@type='submit']")));
		btn_yes.click();
		WebElement btn_clipboard = wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//button[@class='MuiButtonBase-root MuiButton-root MuiButton-contained MuiButton-containedPrimary']")));
		btn_clipboard.click();
		WebToken = (String) Toolkit.getDefaultToolkit().getSystemClipboard().getData(DataFlavor.stringFlavor);
		if (WebToken.contains("eY")) {
			System.out.println("Successfully extracted Web Authentication Token");
		}
		driver.quit();
	}
}