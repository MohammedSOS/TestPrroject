package Day1;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class FirstSeleniumScript {

	public static void main(String[] args) throws InterruptedException {
	System.setProperty("webdriver.chrome.driver", "C:\\seleniumtraining\\chromedriver.exe");
	WebDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.manage().deleteAllCookies();
	driver.manage().timeouts().pageLoadTimeout(40,TimeUnit.SECONDS);
	driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	driver.get("https://www.facebook.com/");
	driver.findElement(By.name("email")).sendKeys("achakoussama7@gmail.com");
	driver.findElement(By.name("pass")).sendKeys("123456789abcd");
	driver.findElement(By.id("u_0_b")).click();
	/*Thread.sleep(10000);
	driver.close();*/
   
	}

}
