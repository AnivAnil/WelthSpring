package WelthSpring;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class NewTest {

	private static WebDriver driver;
    private WebDriverWait wait;
    private final String filePath = "D:\\test1\\CustomerNames.xls";

    @BeforeMethod
    public void browserlaunch() throws IOException {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
    }

    @Test()
    public void loginpage() throws InterruptedException, IOException {
        driver.get("https://wealthspring.my-portfolio.co.in/app/#/login");
        Thread.sleep(1000);
        WebElement username = driver.findElement(By.xpath("//input[@name='email']"));
        username.sendKeys("WealthsAdmin");
        Thread.sleep(1000);
        WebElement password = driver.findElement(By.xpath("//input[@name='password']"));
        password.sendKeys("aditya90");
        Thread.sleep(1000);
        WebElement loginbutton = driver.findElement(By.xpath("//input[@type='submit']"));
        loginbutton.click();
        Thread.sleep(1000);
        // String filePath = "D:\\test1\\CustomerNames.xls";
        FileInputStream fileInputStream = new FileInputStream(filePath);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet sheet = hssfWorkbook.getSheet("sheet1");
        int lastRowNum = sheet.getLastRowNum();

        // Read Excel data only once
        List<String> clientNames = new ArrayList<>();
        for (int i = 1; i <= lastRowNum; i++) {
            HSSFRow row = sheet.getRow(i);
            String stringCellValue = row.getCell(0).getStringCellValue();
            clientNames.add(stringCellValue);
        }

        for (String stringCellValue : clientNames) {
            Thread.sleep(1000);

            WebElement ClientSearch = driver.findElement(By.xpath("//input[@id='selectInput']"));
            ClientSearch.sendKeys(stringCellValue);
            Thread.sleep(1000);
            ClientSearch.sendKeys(Keys.ENTER);
            Thread.sleep(1000);
            WebElement report = driver.findElement(By.xpath("//a[text()='Go to Report']"));
            report.click();
            Thread.sleep(1000);
            String child = driver.getWindowHandle();
            Set<String> Parent = driver.getWindowHandles();
            List<String> list = new ArrayList<String>();
            list.addAll(Parent);
            driver.switchTo().window(list.get(1));
            Thread.sleep(1000);
            WebElement excel_button = driver.findElement(By.xpath("(//span[@class='radioDot'])[3]"));
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", excel_button);
            Thread.sleep(1000);
            WebElement client_wise = driver.findElement(By.xpath("(//span[@class='radioDot'])[9]"));
            client_wise.click();
            Thread.sleep(1000);
            JavascriptExecutor je = (JavascriptExecutor) driver;
            je.executeScript("window.scrollBy(0,300)");
            Thread.sleep(1000);
            WebElement apply_button = driver.findElement(By.xpath("//button[text()='Apply']"));
            apply_button.click();
            Thread.sleep(1000);
            WebElement Portfolio_Performance = driver.findElement(By.xpath("//a[text()='Portfolio Performance']"));
            Portfolio_Performance.click();
            Thread.sleep(1000);
            je.executeScript("window.scrollBy(0,-500)");
            Thread.sleep(1000);
            WebElement Excel2 = driver.findElement(By.xpath("(//span[@class='radioDot'])[3]"));
            Excel2.click();
            Thread.sleep(1000);
            je.executeScript("window.scrollBy(0,300)");
            Thread.sleep(1000);
            WebElement apply_button2 = driver.findElement(By.xpath("//button[text()='Apply']"));
            apply_button2.click();
            Thread.sleep(1000);
            je.executeScript("window.scrollBy(0,-500)");
            Thread.sleep(1000);
            WebElement Capital_Gain_Realized = driver.findElement(By.xpath("//a[text()='Capital Gain Realized']"));
            Capital_Gain_Realized.click();
            Thread.sleep(1000);
            WebElement Excel3 = driver.findElement(By.xpath("(//span[@class='radioDot'])[3]"));
            Excel3.click();
            Thread.sleep(1000);
            je.executeScript("window.scrollBy(0,200)");
            Thread.sleep(1000);
            WebElement apply_button3 = driver.findElement(By.xpath("//button[text()='Apply']"));
            apply_button3.click();
            Thread.sleep(1000);
            driver.close();

            driver.switchTo().window(list.get(0));
            Thread.sleep(1000);
            je.executeScript("window.scrollBy(0,200)");

            WebElement Client = driver.findElement(By.xpath("(//input[@class='hiddenInputFld'])[1]"));
            // Client.click();
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", Client);
            Thread.sleep(3000);

        }
        hssfWorkbook.close();
        fileInputStream.close();
        driver.quit();

    }

    @AfterMethod
    public void browserclose() {
        // driver.quit();
        // System.out.println("completed");
    }
}