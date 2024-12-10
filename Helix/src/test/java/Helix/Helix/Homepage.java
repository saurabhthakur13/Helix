package Helix.Helix;

import io.github.bonigarcia.wdm.WebDriverManager;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.function.Function;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.JavascriptExecutor;
import java.util.List;

public class Homepage {

	
	WebDriver driver;
	Actions hover;
	Workbook workbook;
	Sheet sheet;
	int rowNum = 1; // Start from the second row for test results.

	@BeforeTest
	void openbrowser() throws IOException {
		// Initialize WebDriver
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.get("https://helix-watches.com/");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		hover = new Actions(driver);

		// Initialize Excel workbook and sheet
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("Test Results");

		// Create headers in Excel
		Row headerRow = sheet.createRow(0);
		headerRow.createCell(0).setCellValue("Test Case");
		headerRow.createCell(1).setCellValue("Status");
		headerRow.createCell(2).setCellValue("Error Message");
	}

	@Test
	void TC_01() {
		try {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement headerlogo = wait
					.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@class='img-regular']")));
			headerlogo.click();
			logResult("TC_01", "Pass", null);
		} catch (Exception e) {
			logResult("TC_01", "Fail", e.getMessage());
			Assert.fail("Error in TC_01: " + e.getMessage(), e); // Force test failure
		}
	}

	@Test
	void TC_02() {
		try {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
			WebElement headermenunew = wait.until(ExpectedConditions.elementToBeClickable(By.id("new")));
			hover.moveToElement(headermenunew).perform();

			WebElement insidemenu1 = wait.until(ExpectedConditions.elementToBeClickable(
					By.xpath("//img[@src ='//helix-watches.com/cdn/shop/files/helix-men.png?v=1721233356']")));
			insidemenu1.click();
			driver.navigate().back();

			hover.moveToElement(headermenunew).perform();
			WebElement insidemenu2 = wait.until(ExpectedConditions.elementToBeClickable(
					By.xpath("//img[@src ='//helix-watches.com/cdn/shop/files/awomen-rrival.png?v=1721233356']")));
			insidemenu2.click();
			driver.navigate().back();

			logResult("TC_02", "Pass", null);
		} catch (Exception e) {
			logResult("TC_02", "Fail", e.getMessage());
			Assert.fail("Error in TC_02: " + e.getMessage(), e); // Force test failure
		}
	}

	@Test
	void TC_03() {
		try {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

			String[] submenusinsidemen = { "(//a[@href='/collections/trending'])[1]",
					"(//a[@href = '/collections/exclusive-men'])[1]", "(//a[@href = '/collections/fashion-men'])[1]",
					"(//a[@href = '/collections/classics-men'])[1]",
					"(//a[@href = '/collections/shop-by-watch-type-men'])[1]",
					"(//a[@href = '/collections/analog-watches-men'])[1]",
					"(//a[@href = '/collections/digital-watches'])[1]",
					"(//a[@href = '/collections/shop-by-dial-shape-men'])[1]", "(//a[@href = '/collections/round'])[1]",
					"(//a[@href = '/collections/square'])[1]", "(//a[@href = '/collections/octagon'])[1]",
					"(//a[@href = '/collections/shop-by-strap-type'])[1]",
					"(//a[@href = '/collections/leather-watches-men'])[1]",
					"(//a[@href = '/collections/resin-men'])[1]",
					"(//a[@href = '/collections/stainless-steel-men'])[1]", };

			for (String submenu : submenusinsidemen) {
				WebElement headermenumen = wait.until(ExpectedConditions.elementToBeClickable(By.id("men")));
				hover.moveToElement(headermenumen).perform();

				WebElement submenuElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(submenu)));
				submenuElement.click();
				driver.navigate().back();
				driver.navigate().refresh();
			}

			logResult("TC_03", "Pass", null);
		} catch (Exception e) {
			logResult("TC_03", "Fail", e.getMessage());
			Assert.fail("Error in TC_03: " + e.getMessage(), e); // Force test failure
		}
	}

	@Test

	void TC_04() {
		try {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

			String[] submenuinsidewomen = { "(//a[@href = '/collections/womens-trending-watches'])[1]",
					"(//a[@href = '/collections/bestseller-women'])[1]",
					"(//a[@href = '/collections/exclusive-women'])[1]",
					"(//a[@href = '/collections/fashion-women'])[1]", "(//a[@href = '/collections/classic-women'])[1]",
					"(//a[@href = '/collections/shop-by-watch-type-women'])[1]",
					"(//a[@href = '/collections/analog-watches-women'])[1]",
					"(//a[@href = '/collections/shop-by-dial-shape-women'])[1]",
					"(//a[@href = '/collections/round-women'])[1]", "(//a[@href = '/collections/square-women'])[1]",
					"(//a[@href = '/collections/shop-by-strap-type-women'])[1]",
					"(//a[@href = '/collections/leather-watches-women'])[1]",
					"(//a[@href = '/collections/brass-women'])[1]",
					"(//a[@href = '/collections/stainless-steel-women'])[1]", };

			for (String submenu : submenuinsidewomen) {
				WebElement headermenuwomen = wait.until(ExpectedConditions.elementToBeClickable(By.id("women")));
				hover.moveToElement(headermenuwomen).perform();

				WebElement submenuelement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(submenu)));
				submenuelement.click();
				driver.navigate().back();
				driver.navigate().refresh();

			}
			logResult("TC_04", "Pass", null);
		} catch (Exception e) {
			logResult("TC_04", "Fail", e.getMessage());
			Assert.fail("Error in TC_04:" + e.getMessage(), e);
		}

	}

	@Test

	void TC_05() {
		try {
			WebDriverWait Wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement headermenuoffer = Wait.until(ExpectedConditions.elementToBeClickable(By.id("offers")));
			headermenuoffer.click();
			logResult("TC_05", "Pass", null);
		} catch (Exception e) {
			logResult("TC_05", "Fail", e.getMessage());
			Assert.fail("Error in TC_05" + e.getMessage(), e);
		}
	}

	@Test

	void TC_06() {
		try {
			WebDriverWait Wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement headermenublog = Wait.until(ExpectedConditions.elementToBeClickable(By.id("blogs")));
			headermenublog.click();
			logResult("TC_06", "Pass", null);
		} catch (Exception e) {
			logResult("TC_06", "Fail", e.getMessage());
			Assert.fail("Error in TC_06" + e.getMessage(), e);

		}
	}

	@Test

	void TC_07() {
		try {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			String[] keywords = { "watch", "smartwatch", "analog watch" };

			for (String keyword : keywords) {
				WebElement searchInput = wait
						.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@placeholder='Search']")));

				searchInput.sendKeys(keyword);
				searchInput.sendKeys(Keys.ENTER);

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@placeholder='Search']")));

				Thread.sleep(5000); // Wait for 10 seconds (5,000 milliseconds)

				System.out.println("Search results displayed for keyword: " + keyword);

				// Go back to the previous page (home page)
				driver.navigate().to("https://helix-watches.com/");
				driver.navigate().refresh();

				// Wait for 5 seconds before the next search to ensure smooth navigation
				Thread.sleep(5000); // Optional, you can adjust this time as needed
			}

			logResult("TC_07", "Pass", null);
		} catch (Exception e) {
			logResult("TC_07", "Fail", e.getMessage());
			Assert.fail("Error in TC_07: " + e.getMessage(), e);
		}
	}

	@Test

	void TC_08() {
		try {
			WebDriverWait Wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement MainBannerClick = Wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
					"//img[@src='//helix-watches.com/cdn/shop/files/1920x822_3256e122-3d8c-4f7a-9f86-43d820ed384b.jpg?v=1721930512']")));
			MainBannerClick.click();
			driver.navigate().to("https://helix-watches.com/");
			logResult("TC_08", "Pass", null);
		} catch (Exception e) {
			logResult("TC_08", "Fail", e.getMessage());
			Assert.fail("Error in Tc_08:" + e.getMessage(), e);

		}
	}

	@Test
	void TC_09() {
	    try {
	        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	        JavascriptExecutor js = (JavascriptExecutor) driver;

	        // Scroll to the 2nd section and interact with the first part
	        WebElement secondSectionContainer = wait.until(ExpectedConditions
	                .presenceOfElementLocated(By.xpath("//section[@class='yellow-black padding-tb pb0']")));
	        js.executeScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'start' });", secondSectionContainer);

	        WebElement firstDivElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
	                "//img[@src='//helix-watches.com/cdn/shop/files/686x686-Men_f0ea4e9b-7615-48de-b2d0-519f14d2442d_686x.jpg?v=1721930557']")));
	        js.executeScript("arguments[0].click();", firstDivElement);

	        // Navigate back, refresh, and relocate the section
	        driver.navigate().back();
	        driver.navigate().refresh();

	        // Relocate the section and the second part element
	        secondSectionContainer = wait.until(ExpectedConditions
	                .presenceOfElementLocated(By.xpath("//section[@class='yellow-black padding-tb pb0']")));
	        js.executeScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'start' });", secondSectionContainer);

	        WebElement secondDivElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
	                "//img[@src='//helix-watches.com/cdn/shop/files/686x686-women_98f1b0cf-6749-4d5e-8370-37495d2ea743_686x.jpg?v=1721930579']")));
	        js.executeScript("arguments[0].click();", secondDivElement);

	        logResult("TC_09", "Pass", null);
	    } catch (Exception e) {
	        logResult("TC_09", "Fail", e.getMessage());
	        Assert.fail("Error in TC_09: " + e.getMessage(), e);
	    }
	}

	

	@AfterTest
	void closeBrowser() throws IOException {
		// Save the workbook to a specific folder
		File file = new File("D:\\TestResults\\TestResults.xlsx");

		// Ensure the directory exists, create it if not
		file.getParentFile().mkdirs();

		FileOutputStream fileOut = new FileOutputStream(file);
		workbook.write(fileOut);
		fileOut.close();
		workbook.close();

		// Close the browser
		driver.quit();
	}

	// Log test results
	void logResult(String testCase, String status, String errorMessage) {
		Row row = sheet.createRow(rowNum++);
		row.createCell(0).setCellValue(testCase);
		row.createCell(1).setCellValue(status);
		row.createCell(2).setCellValue(errorMessage != null ? errorMessage : "None");
	}
}
