package com.clover.pageobjects;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;


public class HomePage {

	WebDriver driver;

	public HomePage(WebDriver wdriver) {
		driver = wdriver;
		PageFactory.initElements(wdriver, this);
	}

	public HomePage() {
		// TODO Auto-generated constructor stub
	}

	@FindBy(xpath = "//input[@title='Search']")
	WebElement searchTextField;

	@FindBy(xpath = "//div[@aria-label='Search by voice']//*[name()='svg']")
	WebElement searchBtn;

	public void ClearSearchTextField() {
		searchTextField.clear();
	}

	public void enterText(String txt) {
		searchTextField.sendKeys(txt);
	}

	public void searchBtn() {
		searchBtn.click();
	
	}

	

}








