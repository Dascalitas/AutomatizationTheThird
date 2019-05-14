package com.dascalitas;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class googleSearch {
    public static void main(String[] args) throws InterruptedException {
        WebDriver driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        driver.manage().timeouts().pageLoadTimeout(300, TimeUnit.SECONDS);

        Workbook wb = new HSSFWorkbook();

        driver.get("http://www.google.com");

        WebElement element = driver.findElement(By.name("q"));

        element.sendKeys("innuendo");

        element.submit();

//        (new WebDriverWait(driver, 10)).until(new ExpectedCondition<Boolean>() {
//            public Boolean apply(WebDriver d) {
//                return d.getTitle().toLowerCase().startsWith("cheetah");
//            }
//        });

        Sheet sheet = wb.createSheet(driver.getTitle());

        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("Site Name");
        row.createCell(1).setCellValue("Opened URL");
        row.createCell(2).setCellValue("Number of Occurrences of Searched Word");

        List<WebElement> links = driver.findElements(By.xpath("//div[@class='r']/a"));

        for(int link=0;link<links.size();link++){
            List<WebElement> links2 = driver.findElements(By.xpath("//div[@class='r']/a"));
            links2.get(link).click();

            Row row2 = sheet.createRow(link + 1);
            row2.createCell(0).setCellValue(driver.getTitle());
            row2.createCell(1).setCellValue(driver.getCurrentUrl());
            row2.createCell(2).setCellValue(StringUtils.countMatches(driver.getPageSource().toLowerCase(),"innuendo"));
            driver.navigate().back();
            Thread.sleep(3000);
        }
        driver.quit();

        try (OutputStream fileOut = new FileOutputStream("Innuendo.xls")) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
