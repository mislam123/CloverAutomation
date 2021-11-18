package com.clover.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	Logger log = Logger.getLogger(BaseClass.class.getName());
	Properties prop = new Properties();
	WebDriver driver;

	@BeforeTest
	public  void LaunchApp() throws Exception{
		
		
			File src = new File("Configuration/config.properties");
			prop = new Properties();
			FileInputStream fis = new FileInputStream(src);
			prop.load(fis);
	
		String browserName = prop.getProperty("browser");
		
		if(browserName.equalsIgnoreCase("chrome")){
			WebDriverManager.chromedriver().setup();
			driver=new ChromeDriver();
		}
		else if(browserName.equalsIgnoreCase("firefox")){
			WebDriverManager.firefoxdriver().setup();
			driver = new FirefoxDriver();
		}
		else if(browserName.equalsIgnoreCase("ie")){
			WebDriverManager.iedriver().setup();
			driver = new InternetExplorerDriver();
		}
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("ttps://www.google.com/");
		//driver.get(prop.getProperty("baseURL"));
		
	}
	@AfterTest
	public void StopApp() throws Exception {
		driver.quit();
		log.info("~~~~~~~~~~~~~ APP CLOSED ~~~~~~~~~~~~~~~");
		log.info("");
		log.info("");
		log.info("");
	}

	public void clickByXpath(String xpath) throws Exception {
		driver.findElement(By.xpath(xpath)).click();
	}

	public void typeByXpath(String xpath, String text) throws Exception {
		driver.findElement(By.xpath(xpath)).sendKeys(text);
	}

	public void clickByName(String name) throws Exception {
		driver.findElement(By.name(name)).click();
	}


}


