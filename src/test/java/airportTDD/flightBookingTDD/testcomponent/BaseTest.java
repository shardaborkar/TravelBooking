package airportTDD.flightBookingTDD.testcomponent;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;

//import com.aventstack.extentreports.ExtentTest;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import airportTDD.pageComponents.SearchFlight;
import io.github.bonigarcia.wdm.WebDriverManager;
import mainPageObject.TravelHomePage;


public class BaseTest {
    public WebDriver driver;
	public SearchFlight searchflight;
	public XSSFSheet sheet;
	int count=1;
	public TravelHomePage travelHomePage;
	public WebDriver initializeDriver() throws IOException{
		
		
		Properties pro=new Properties();
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+ "\\src\\main\\java\\resources\\global.properties");
		pro.load(fis);
		
		String browserName=System.getProperty("browser")!=null?System.getProperty("browser") :pro.getProperty("browser");
		if(browserName.contains("chrome")){
	    ChromeOptions options=new ChromeOptions();
	    options.addArguments("--remote-allow-origins=*");
		WebDriverManager.chromedriver().setup();
		if(browserName.contains("headless")) {
		options.addArguments("headless");
		}
		driver=new ChromeDriver(options);
		
		}
		
		else if(browserName.equalsIgnoreCase("firefox")) {
			FirefoxOptions options=new FirefoxOptions();
			
	       WebDriverManager.firefoxdriver().setup();
	       if(browserName.contains("headless")) {
	   		options.addArguments("headless");
	   		}
		   driver=new FirefoxDriver(options);
		}
		
		else if(browserName.equalsIgnoreCase("edge")) {
			EdgeOptions options=new EdgeOptions();
			WebDriverManager.edgedriver().setup();
			options.addArguments("--remote-allow-origins=*");
	       if(browserName.contains("headless")) {
	   		options.addArguments("headless");
	   		}
		   driver=new EdgeDriver(options);
		}
		/*
		DesiredCapabilities dr = new DesiredCapabilities();
	        dr.setBrowserName("chrome"); 
                driver= new RemoteWebDriver(new java.net.URL("http://20.48.204.0:4444⁠"),dr);
		*/
		driver.manage().window().maximize();
		driver.manage().window().setSize(new Dimension(1440,900));
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		return driver;
	}
	
	public List<HashMap<String, String>> getJSONdata(String filePath) throws IOException {
		
		String jsonContent=FileUtils.readFileToString(new File(filePath),StandardCharsets.UTF_8);
		ObjectMapper mapper=new ObjectMapper();
		List<HashMap<String,String>> data=mapper.readValue(jsonContent, new TypeReference<List<HashMap<String,String>>>(){});
		
		return data;
	}	
	
	public String getScreenshot(String testCaseName,WebDriver driver) throws IOException {
		TakesScreenshot s=(TakesScreenshot)driver;
		File file=s.getScreenshotAs(OutputType.FILE);
		File destFile=new File(System.getProperty("user.dir") + "//screenshots//"+testCaseName+".png");
		FileUtils.copyFile(file, destFile);
		return System.getProperty("user.dir") + "//screenshots//"+testCaseName+".png";
	}
	
	public Object[][] getExcelData(String filepath,String sheetName) throws IOException
	 {
	 //fileInputStream argument
		 
	 File file=new File(filepath);
	 FileInputStream fis=new FileInputStream(file);
	 XSSFWorkbook workbook=new XSSFWorkbook(fis);

	 int sheets=workbook.getNumberOfSheets();
	 for(int i=0;i<sheets;i++)
	 {
	 if(workbook.getSheetName(i).equalsIgnoreCase(sheetName))
	 {
	 sheet=workbook.getSheetAt(i);
	 break;
	 }}
	 //Identify Testcases coloum by scanning the entire 1st row
	 int rownum=sheet.getLastRowNum();
	 
	 int column=sheet.getRow(0).getPhysicalNumberOfCells();
	 int actualColumn=column-1;
	 Object[][] obj=new Object[rownum][actualColumn];
	 for(int k=1;k<=rownum;k++) {
		 for(int j=1;j<column;j++) {
		obj[k-1][j-1]=sheet.getRow(k).getCell(j).getStringCellValue();
		
	 }}
	 workbook.close();
	 return obj;
	 }
	
	
    
	@BeforeMethod(alwaysRun=true)
	public void launchApp() throws IOException {
		driver=initializeDriver();
		travelHomePage=new TravelHomePage(driver);
//		searchflight=new SearchFlight(driver);
//		searchflight.goTo();
		travelHomePage.goTo();
//		return travelHomePage;
		
	}
  @AfterMethod(alwaysRun=true)
	  public void tearDown() {
	  driver.close();
  }
  }

