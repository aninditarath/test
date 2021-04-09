package demo;

import java.time.Instant;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Demo {
	public static WebDriver driver = null;

	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws InterruptedException {

		System.setProperty("webdriver.chrome.driver",
				System.getProperty("user.dir")
						+ "//resources//chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.worldometers.info/world-population/");
		driver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
		long endTime = System.currentTimeMillis() + 15000;
		while (System.currentTimeMillis() < endTime) {
			String current_pupulation = driver.findElement(
					By.xpath("//div[@class='maincounter-number']")).getText();
			
			String birthstoday=driver.findElement(By.xpath("//div[@class='col1in']//div[2]//div[2]")).getText();
			String birthsthisyear=driver.findElement(By.xpath("//div[@class='col2in']//div[2]//div[2]")).getText();
			System.out.println("Current population : "+current_pupulation + "  and , "+ "Birthstoday : "+birthstoday+"  and , "+"Birthsthisyear : "+birthsthisyear);
		}

		driver.quit();
	}

}
