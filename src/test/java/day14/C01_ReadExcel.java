package day14;

import org.apache.poi.ss.usermodel.*;
import org.junit.Assert;
import org.junit.Test;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class C01_ReadExcel {


    @Test
    public void readExcelTest1() throws IOException {

        //- Dosya yolunu bir String degiskene atayalim
        String dosyaYolu = "src/resources/ulkeler.xlsx";

        //- FileInputStream objesi olusturup,parametre olarak dosya yolunu girelim
        FileInputStream fis = new FileInputStream(dosyaYolu); //Olusturmuş olduğumuz dosyayı sistemde işleme alır

        //- Workbook objesi olusturalim,parameter olarak fileInputStream objesini girelim
        //- WorkbookFactory.create(fileInputStream)
        Workbook workbook = WorkbookFactory.create(fis); //Workbook objesiyle fis oblesi ile akışa aldığımız dosyamızla bir excell dosyası create ettik

        //- Sheet objesi olusturun workbook.getSheetAt(index)
        Sheet sheet = workbook.getSheet("Sayfa1"); //Excel dosyamızda çalışmak istediğimiz sayfayı bu şekilde seçeriz

        //- Row objesi olusturun sheet.getRow(index)
        Row row = sheet.getRow(3); // Sayfa bir deki 3. satırı bu şekilde seçeriz

        //- Cell objesi olusturun row.getCell(index)
        Cell cell = row.getCell(3); // Satır seçimi yapıldıktan sonra hücre seçimi bu şekilde yapılır
        System.out.println(cell);

        //- 3. index'deki satirin 3. index'indeki datanin Cezayir oldugunu test edin
        String expectedData = "Cezayir";
        String actualData = cell.toString();
        Assert.assertEquals(expectedData,actualData);
        Assert.assertEquals(expectedData,cell.getStringCellValue());   // 38. satirin 2. yolu

    }



    /*
    Ara-->dosyaYolu
    Windows Gezgini ile ac-->FileInputStream
    Excel i ac-->Workbook
    Sayfa1 e git-->Sheet
    Satiri bul-->Row
     */

    /*
    Exceldeki bir dosyaya ve dosyadaki bir bölüme ulaşmak istediğimizde
    uygulayacağımız aşamalar :
    Ara-->dosyaYolu      --> String dosyaYolu="src/resources/ulkeler.xlsx";
    Windows Gezgini ile ac- ->FileInputStream    --> FileInputStream fis = new FileInputStream(dosyaYolu);
    Excel i ac              -->Workbook        --> Workbook workbook = WorkbookFactory.create(fis)

    workbook olusturduktan sonra :
    String actualData = workbook.
                getSheet("Sayfa1")
                .getRow(3)
                .getCell(3)
                .toString();

     Sayfa1 e git-->Sheet    --> getSheet()
     Satiri bul-->Row       --> getRow()
     Sutunu bul-->Cell     --> getCell()
     Secilen bolumu getir  --> toString()
     */

}
