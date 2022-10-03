package day12;

import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import utilities.TestBaseBeforeAfter;

public class Odev1 extends TestBaseBeforeAfter {


    /*
   1."http://webdriveruniversity.com/Actions"sayfasina gidin
   2."Hover over Me First" kutusunun ustune gelin
   3."Link 1" e tiklayin
   4.Popup mesajini yazdirin
   5.Popup'i tamam diyerek kapatin
   6."Click and hold" kutusuna basili tutun
   7."Click and hold" kutusunda cikan yaziyi yazdirin
   8."Double click me" butonunu cift tiklayin
     */


    @Test
    public void test() {

        //1."http://webdriveruniversity.com/Actions" sayfasina gidin
        driver.get("http://webdriveruniversity.com/Actions");


        // 2."Hover over Me First" kutusunun ustune gelin
        Actions actions = new Actions(driver);
        WebElement hoverOver = driver.findElement(By.xpath("//*[text()='Hover Over Me First!']"));
        actions.moveToElement(hoverOver).perform();


        //3.Link 1" e tiklayin
        driver.findElement(By.xpath("(//*[text()='Link 1'])[1]")).click();


        //4.Popup mesajini yazdirin
        System.out.println(driver.switchTo().alert().getText());


        //5.Popup'i tamam diyerek kapatin
        driver.switchTo().alert().accept();


        //6."Click and hold" kutusuna basili tutun
        WebElement clickAndHoldBut = driver.findElement(By.id("click-box"));
        actions.clickAndHold(clickAndHoldBut).perform();


        //7."Click and hold" kutusunda cikan yaziyi yazdirin
        System.out.println(clickAndHoldBut.getText());


        //8."Double click me" butonunu cift tiklayin
        WebElement doubleClickMe = driver.findElement(By.xpath("//*[text()='Double Click Me!']"));
        actions.doubleClick(doubleClickMe).perform();


    }

}
