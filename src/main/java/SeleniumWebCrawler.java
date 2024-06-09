import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class SeleniumWebCrawler {

//    爬取网页的所有a连接
    static List<String> urls = new ArrayList<String>();
    public static void main(String[] args) {

        Workbook workbook = new XSSFWorkbook();
        // 创建一个工作表
        Sheet sheet = workbook.createSheet("NewsData");

//设置系统属性指定谷歌驱动地址
        System.setProperty("webdriver.chrome.driver", "D:\\chromedriver_win32\\chromedriver.exe");

        // 创建ChromeDriver对象
        WebDriver driver =  new ChromeDriver();

        try {
            // 打开目标网页
            driver.get("https://xz.chsi.com.cn/occucase/index.action");

            try {
                Thread.sleep(3000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            Actions action = new Actions(driver);



//
// 鼠标右键点击指定的元素
            List<WebElement> btns = driver.findElements(By.className("ivu-page-item"));

//            System.out.println(pageSource);
            int a = 1;
//            获取13页的数据的a链接
            while(a<66){
                getAs(driver,a);/**/
                //            点击下一个页面
                if (a==65){
                    action.click(btns.get(0)).perform();
                }else {
                    action.click(btns.get(1)).perform();
                }
                action.release().perform();
                try {
                    Thread.sleep(1000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                a++;
            }

            getText(driver,sheet,workbook);
        } finally {
            // 关闭浏览器驱动
            driver.quit();
        }
    }



    public static void getText(WebDriver driver,Sheet sheet,Workbook workbook){
        int i =0 ;
        for (String url : urls) {
            int j = 0;
            driver.get(url);
            if (i==0){
                WebElement input1 = driver.findElement(By.id("username"));
                WebElement input2 = driver.findElement(By.id("password"));
                input1.sendKeys("account");
                input2.sendKeys("password");
                WebElement btn = driver.findElement(By.className("login-btn"));
                Actions action = new Actions(driver);
                action.click(btn).perform();
                i++;
            }

            try {
                Thread.sleep(1000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
//            String title = driver.findElement(By.className("title")).getText();
            // 创建一行
            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
//          1，8，11
            List<WebElement> div1 = driver.findElements(By.className("detail-des"));
//                System.out.println(div1.get(0).getText());
//           其余都在这里
            List<WebElement> div2 = driver.findElements(By.className("detail-text"));
            for (WebElement webElement : div1) {
                if(webElement.getText().equals("")){
                    continue;
                }
                Cell cellDes= row.createCell(j);
                cellDes.setCellValue(webElement.getText());
                j++;
            }

            for (WebElement webElement : div2) {
                if(webElement.getText().equals("")){
                    continue;
                }

                Cell cellDes= row.createCell(j);
                cellDes.setCellValue(webElement.getText());
                j++;
            }


                try (FileOutputStream fileOut = new FileOutputStream("news_data.xlsx")) {
                    workbook.write(fileOut);
                    System.out.println("Data written to Excel file successfully.");
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
        }



public static  void getAs(WebDriver driver,int page){
//获取a连接的链接
    WebElement element = driver.findElement(By.className("xz-zyrw-list"));
    List<WebElement> a = element.findElements(By.tagName("a"));
    // 获取网页源代码
            String pageSource = driver.getPageSource();
    System.out.println("第"+page+"页");
    for (WebElement as : a) {
        if (as.getAttribute("href")==null){
            continue;
        }
        urls.add(as.getAttribute("href"));
        System.out.println(as.getAttribute("href"));

    }
}
}

