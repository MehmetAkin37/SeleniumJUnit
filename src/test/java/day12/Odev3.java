package day12;

import org.junit.Assert;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.interactions.Actions;
import utilities.TestBaseBeforeAfter;

import java.util.ArrayList;
import java.util.List;

public class Odev3  extends TestBaseBeforeAfter {

    /*
    1."http://webdriveruniversity.com/"adresinegidin
    2."Login Portal" a kadar asagiinin
    3."Login Portal" a tiklayin 4.Diger window'agecin
    5."username" ve "password" kutularina degeryazdirin
    6."login" butonunabasin
    7.Popup'ta cikan yazinin "validation failed" oldugunu test edin 8.Ok diyerek Popup'ikapatin
    9.Ilk sayfaya geri donun
    10.Ilk sayfaya donuldugunu testedin
     */


    @Test
    public void test1(){

        //1."http://webdriveruniversity.com/" adresine gidin
        driver.get("http://webdriveruniversity.com/");


        //2."Login Portal" a  kadar asagi inin
        Actions actions = new Actions(driver);
        actions.sendKeys(Keys.PAGE_DOWN).perform();


        //3."Login Portal" a tiklayin
        driver.findElement(By.xpath("//*[text()='LOGIN PORTAL']")).click();


        // 4.Diger window'a gecin
        List<String> windowHandle = new ArrayList<String>(driver.getWindowHandles());
        driver.switchTo().window(windowHandle.get(1));


        //5."username" ve  "password" kutularina deger yazdirin
        driver.findElement(By.xpath("//*[@id='text']")).sendKeys("M.Akin");
        driver.findElement(By.xpath("//*[@id='password']")).sendKeys("154614");


        //6."login" butonuna basin
        driver.findElement(By.xpath("//*[@id='login-button']")).click();


        //7.Popup'ta cikan yazinin "validation failed" oldugunu test edin
        String expectedPopYazi = "validation failed";
        String actualPopYazi =driver.switchTo().alert().getText();
        Assert.assertEquals(actualPopYazi, expectedPopYazi);


        //8.Ok diyerek Popup'i kapatin
        driver.switchTo().alert().accept();


        //9.Ilk sayfaya geri donun
        driver.switchTo().window(windowHandle.get(0));


        //10.Ilk sayfaya donuldugunu test edin
        String expectedTitleIlksyf ="WebDriverUniversity.com";
        String actualTitleIlksyf =driver.getTitle();
        Assert.assertEquals(actualTitleIlksyf, expectedTitleIlksyf);

    }


}
