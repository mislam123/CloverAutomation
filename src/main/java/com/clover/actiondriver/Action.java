package com.clover.actiondriver;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

import com.clover.base.BaseClass;

public class Action extends BaseClass{
	
	public static void scrollByVisibilityOfElement(WebDriver ldriver, WebElement element){
		//JavascriptExecutor js = (JavascriptExecutor) driver;
		//js.executeScript("arguments[0].scrollIntoView();", element);
	}
	public static void clickByMoveToElement(WebDriver ldriver, WebElement locatorName){
		Actions act = new Actions (ldriver);
		act.moveToElement(locatorName).click().build().perform();
	}
	
	public static void findElement(WebDriver ldriver, WebElement element){
		boolean flag = false;
		try{
			element.isDisplayed();
			flag = true;
		}catch(Exception e){
			flag = false;
		}finally{
			if(flag){
				System.out.println("Successfully found the element at");
			}else{
				System.out.println("Unable to locate the element at");
			}
		}
}

}