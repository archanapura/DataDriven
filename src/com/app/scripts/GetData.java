package com.app.scripts;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import com.app.libraries.ExcelLib;

public class GetData 
{
	public static void main(String[] args) throws Exception{

		System.out.println("Execution Started");
		System.setProperty("webdriver.chrome.driver", "./drivers/chromedriver.exe");
		WebDriver driver =new ChromeDriver();
		ExcelLib lib =new ExcelLib();
		driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
		driver.get("https://gmail.com");
		System.out.println(driver.getTitle());
		int rowno=lib.TotalRows("./Data/Book1.xlsx","Sheet1");
		System.out.println(lib.TotalCells("./Data/Book1.xlsx","Sheet1",0));
		/*  try 
		   {*/
		if(driver.getTitle().equals("Gmail"))
		{
			for(int i=1;i<=rowno;i++)
			{
				String usrname= lib.readData("./Data/Book1.xlsx","Sheet1",i, 0);
				System.out.println(usrname);
				Thread.sleep(3000);
				Actions act = new Actions(driver);
				// act.moveToElement(driver.findElement(By.xpath("//div[text()='Email or phone']"))).click().build().perform();
				//  act.sendKeys(usrname).build().perform();
				Thread.sleep(3000);
				driver.findElement(By.name("identifier")).sendKeys(usrname);
				driver.findElement(By.xpath("//span[text()='Next']")).click();
				Thread.sleep(3000);
				String pasword= lib.readData("./Data/Book1.xlsx","Sheet1",i, 1);
				System.out.println(pasword);
				Thread.sleep(3000);
				act.moveToElement(driver.findElement(By.name("password"))).click().perform();
				act.sendKeys(pasword).perform();
				Thread.sleep(3000);
				driver.findElement(By.xpath("//span[text()='Next']")).click();
				
				driver.navigate().back();
				Thread.sleep(3000);
				driver.findElement(By.name("identifier")).clear();
			}
		}
		/* }catch(Exception e) 
		   {
			   System.out.println("loginpage not displayed");
		   }*/
		System.out.println("Execution Ended");



		/*System.out.println(lib.readData("./Data/Book1.xlsx", "Sheet1", 0, 0));
	lib.writeData("./Data/Book1.xlsx", "Sheet1", 1, 4, "Archana");*/

	}


}
