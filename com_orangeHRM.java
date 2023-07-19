package com.HR.project;

import java.io.FileInputStream;
import java.time.Duration;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;

public class com_orangeHRM {

	public static void main(String[] args) {
		try {
	    	ChromeDriver driver =new ChromeDriver();
		    driver.manage().window().maximize();
		    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		    driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");
		    Thread.sleep(2000);
		    
		   FileInputStream file = new FileInputStream("C:\\excelr_workspace\\hrm.xlsx"); 
		    XSSFWorkbook workbook = new XSSFWorkbook(file);
		    XSSFSheet sheet = workbook.getSheet("Sheet1");
		    int rowcount = sheet.getLastRowNum();
		    int colcount = sheet.getRow(1).getLastCellNum();
		    
		    for(int i=1; i<=rowcount; i++)
		    {
		    	String A= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(1));
		    	String B= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(2));
		    	String C= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(3));
		    	String D= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(4));
		    	String E= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(5));
		    	String F= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(6));
		    	String G= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(7));
		    	String H= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(8));
		    	String I= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(9));
		    	String J= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(10));
		    	String K= new DataFormatter().formatCellValue(sheet.getRow(i).getCell(11));
		    	
		        driver.findElement(By.name("username")).sendKeys("Admin");
		        Thread.sleep(500);
		        driver.findElement(By.name("password")).sendKeys("admin123");
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[3]/button")).click();
		        Thread.sleep(500);
		        
		    //Time at Work
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/div[1]/div[2]/button")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/div/form/div[2]/div/div/div/div[2]/textarea")).sendKeys(A);
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/div/form/div[3]/button")).click();
		        Thread.sleep(1000);
		        
	      //my action 
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[1]/p")).click();
		        Thread.sleep(500);
		        
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[2]/p")).click();		        
		        Thread.sleep(500);
		        
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[3]")).click();
		        Thread.sleep(500);
		        
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[4]/p")).click();
		        Thread.sleep(500);
		        
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
		        
		        
		   //Quick Launch  
		        
	            driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/leave/assignLeave");
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div/div/div[2]/div/div/input")).sendKeys(B);
		        Thread.sleep(500);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/div/div/div[1]")).sendKeys(C);   
		    	Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/form/div[3]/div/div[1]/div/div[2]/div/div/input")).sendKeys("2023-07-25");
		        Thread.sleep(1000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/form/div[3]/div/div[2]/div/div[2]/div/div/input")).sendKeys("2023-07-30");;
		        Thread.sleep(500);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/form/div[4]/div/div/div/div[2]/textarea")).sendKeys(F);
		        Thread.sleep(1000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div/form/div[5]/button")).click();
		        
		        driver.get(G);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
	        	Thread.sleep(3000);
		        driver.get(H);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
				Thread.sleep(3000);
				driver.get(I);
				driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
				Thread.sleep(3000);
				driver.get(J);
				driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
				Thread.sleep(3000);
				driver.get(K);
				Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();
		        
		 //Buzz Latest Posts       
		        
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[4]/div/div[2]/div/div[4]/div/div[2]/p[1]")).click();
		        Thread.sleep(3000);
			    Thread.sleep(1000);
	            driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a")).click();

		 //Employees on Leave Today     
		       driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[5]/div/div[1]/i")).click();
		       Thread.sleep(2000);
		       driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[5]/div[2]/div/div/div/form/div[1]/div/div[2]/div/label/span")).click();
		       Thread.sleep(2000);
		       driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[5]/div[2]/div/div/div/form/div[2]/button[2]")).click();
		        
        //Employee Distribution by Sub Unit    
		        
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[1]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[1]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[2]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[2]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[3]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[3]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[4]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[4]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[5]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[5]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[6]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[6]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[7]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[6]/div/div[2]/div/ul/li[7]/span[2]")).click();
		        
		        
		  //Employee Distribution by Location      
		
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[1]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[1]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[2]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[2]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[3]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[3]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[4]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[4]/span[2]")).click();
		        Thread.sleep(2000);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[5]/span[2]")).click();
		        Thread.sleep(1200);
		        driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[2]/div/div[7]/div/div[2]/div/ul/li[5]/span[2]")).click();
		        Thread.sleep(2000);
		    }
		   driver.quit();
		    workbook.close();
     	} 
	    catch (Exception e) {
		   System.out.println(e.getMessage());
     	}
  }
}