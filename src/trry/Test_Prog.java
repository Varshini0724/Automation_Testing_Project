package trry;

import java.io.File;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;

@Test
public class Test_Prog {
    static WebDriver driver;
    static WebDriverWait wait;
    static WritableWorkbook outputWorkbook;
    static WritableSheet outputSheet;
    static int outputRow = 0;
    static Sheet inputSheet;

    @BeforeTest
    public void setup() throws Exception {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\vaish\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://www.beyoung.in/");
        wait = new WebDriverWait(driver, Duration.ofSeconds(60));

        // Read input Excel file
        Workbook inputWorkbook = Workbook.getWorkbook(new File("TestInput.xls"));
        inputSheet = inputWorkbook.getSheet(0);

        // Create output Excel workbook and sheet
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        outputWorkbook = Workbook.createWorkbook(new File("TestResults_" + timestamp + ".xls"));
        outputSheet = outputWorkbook.createSheet("Test Results", 0);

        // Add headers to output sheet
        outputSheet.addCell(new Label(0, outputRow, "Step"));
        outputSheet.addCell(new Label(1, outputRow, "Status"));
        outputSheet.addCell(new Label(2, outputRow, "Timestamp"));
        outputRow++;

        writeToExcel("Browser opened", "Pass");
    }

    public void login() throws Exception {
        try {
            // Read email from input Excel
            String email = inputSheet.getCell(1, 0).getContents();

            // Wait for the login button to be clickable
            WebElement loginButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("loginBtn")));
            loginButton.click();
            writeToExcel("Clicked on Login", "Pass");

            // Wait for the phone number input field to be present and interactable
            WebElement phoneNumberField = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("login-numbers")));
            phoneNumberField.sendKeys(email);
            writeToExcel("Entered phone number: " + email, "Pass");

            // Wait for the Continue button to be clickable and click it
            WebElement loginWithOTPButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(text(),'Login with OTP')]")));
            loginWithOTPButton.click();
            writeToExcel("Clicked on Continue button", "Pass");

            // Wait for manual OTP entry
            // This part can be automated if you have access to the OTP generation or interception
            waitForOTPEntry();

            // Use a more general XPath to locate the VERIFY button
            WebElement loginSubmitButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(text(),'Verify')]")));
            loginSubmitButton.click();
            writeToExcel("Clicked on VERIFY button", "Pass");

            // Wait for the login process to complete
            waitForLoginToComplete();

            // Read search keyword from input Excel
            String searchKeyword = inputSheet.getCell(1, 1).getContents();

            // Click on the search button
            WebElement searchButton = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("span.search-mobile")));
            searchButton.click();
            writeToExcel("Clicked on search button", "Pass");

            // Wait for the search input field to appear and enter the keyword
            WebElement searchInput = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Search entire store here...']")));
            searchInput.sendKeys(searchKeyword);
            writeToExcel("Entered '" + searchKeyword + "' in search box", "Pass");
            searchInput.sendKeys(Keys.RETURN);

            // Locate the first product by its href attribute
            WebElement firstProductLink = wait.until(ExpectedConditions.presenceOfElementLocated(
                By.xpath("//a[@href='/rich-black-chinos-for-men']")
            ));

            // Scroll the element into view and click it
            scrollToAndClickElement(firstProductLink);
            writeToExcel("Clicked on the first product in search results", "Pass");

            // Select the size 28 and add the product to the cart
            selectSizeAndAddToCart();

        } catch (NoSuchElementException e) {
            writeToExcel("Element not found: " + e.getMessage(), "Fail");
            e.printStackTrace();
        } catch (Exception e) {
            writeToExcel("Exception: " + e.getMessage(), "Fail");
            e.printStackTrace();
        }
    }

    private void selectSizeAndAddToCart() throws Exception {
        // Click on the size 28
        WebElement size28 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//p[@class='sizevalue-main' and text()='28']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", size28);
        writeToExcel("Selected size 28", "Pass");

        // Click on 'Add to Cart' anchor tag
        WebElement addToCartLink = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("//a[contains(text(), 'Add to Cart')]")
        ));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", addToCartLink);
        writeToExcel("Clicked on 'Add to Cart' link", "Pass");
    }

    private void scrollToAndClickElement(WebElement element) {
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", element);
        wait.until(ExpectedConditions.elementToBeClickable(element));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", element);
    }

    private void waitForOTPEntry() throws Exception {
        // Wait for manual OTP entry (or automate OTP entry if possible)
        Thread.sleep(10000);
        writeToExcel("Waited for OTP entry", "Pass");
    }

    private void waitForLoginToComplete() throws Exception {
        // Wait for login to complete
        Thread.sleep(10000);
        writeToExcel("Waited for login process to complete", "Pass");
    }

    @AfterTest
    public void tearDown() throws Exception {
        // Commented out to keep the browser open
        // if (driver != null) {
        //     driver.quit();
        //     writeToExcel("Browser closed", "Pass");
        // }
        
        // Write and close the Excel workbook
        outputWorkbook.write();
        outputWorkbook.close();
    }

    private void writeToExcel(String step, String status) throws Exception {
        outputSheet.addCell(new Label(0, outputRow, step));
        outputSheet.addCell(new Label(1, outputRow, status));
        outputSheet.addCell(new Label(2, outputRow, LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"))));
        outputRow++;
    }
}
