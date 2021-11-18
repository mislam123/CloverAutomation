package com.clover.testcase;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import org.apache.log4j.Logger;
import org.apache.log4j.xml.DOMConfigurator;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;
import com.clover.dataprovider.ExcelIO;
import io.github.bonigarcia.wdm.WebDriverManager;


public class GoogleSearchTest{
	
	Logger log = Logger.getLogger(GoogleSearchTest.class.getName());
	Properties prop = new Properties();
	WebDriver driver;


	@Test(dataProvider = "TestData", dataProviderClass = ExcelIO.class)
	public void GoogleSearchTest(String InputText, String ExpectedText ) throws Exception{
		
		DOMConfigurator.configure("log4j.xml");

		File src = new File("Configuration/config.properties");
		prop = new Properties();
		FileInputStream fis = new FileInputStream(src);
		prop.load(fis);
		
		WebDriverManager.chromedriver().setup();
		driver=new ChromeDriver();

	driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
	driver.manage().window().maximize();
	//driver.get("ttps://www.google.com/");
	driver.get(prop.getProperty("baseURL"));
	
		log.info("Launch the browser and redirected to the URL");
	
	driver.findElement(By.xpath("//input[@title='Search']")).clear();
	driver.findElement(By.xpath("//input[@title='Search']")).sendKeys(InputText);
		log.info("Entered text for searching");
	Thread.sleep(1000);
	//String getText =driver.findElement(By.xpath("//input[@title='Search'][1]")).getText();
	String getText =driver.findElement(By.xpath("//li[@class='sbct'][1]")).getText();
	System.out.println("@@@@@@@@@@@@@@@@@@@@: "+ getText);
		log.info("Found the search result");

		
	if(ExpectedText.equals(getText)) {
		System.out.println("Verification Passed");
		log.info("Test Passed");
	}else {
		System.err.println("Verification FAILED, Expected : "+ ExpectedText  + "  But found :  " + getText);
		log.info("Test FAILED");
	}
	
	Thread.sleep(1000);
	driver.quit();

	}

	}

