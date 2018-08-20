import java.io.IOException;
import java.util.concurrent.TimeUnit;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import java.io.File;
//import org.junit.*;
//import org.junit.Assert;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.Key;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;
import org.testng.Assert;
import org.testng.annotations.*;
/*import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;
//import java.nio.charset.Charset;*/

public class TestForm {

    private static WebDriver driver;
    private static String Chromedriverpath;
    private static String Chromedriverlog;
    private static String SikuliPath;
    private static String baseUrl;
    private static int timeout=5000;

    //private static Workbook workbook;

    //private static WorkbookSettings ws = new WorkbookSettings();

    private DataManagement data=new DataManagement();

    //public static final Charset UTF_8 = Charset.forName("UTF-8");

    @BeforeClass
    public static void setUp() throws Exception {

        Chromedriverpath = "C:\\Users\\fchautem\\chromedriver.exe";
        Chromedriverlog = "C:\\Users\\fchautem\\chromedriver.log";
        SikuliPath = "D:\\Sikuli\\TestForm\\";

        baseUrl = "http://cthulhu/public/Inscription.nsf/Inscription.xsp";
        System.setProperty("webdriver.chrome.driver", Chromedriverpath);
        System.setProperty("webdriver.chrome.logfile", Chromedriverlog);

        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.manage().timeouts().implicitlyWait(timeout,TimeUnit.MILLISECONDS);
        //driver = new ChromeDriver();
        //java.nio.charset.StandardCharsets.UTF_8.encode(string);
    }

    /*public void loadXLS(String pathname) {

        workbook=null;
        try {
            workbook = Workbook.getWorkbook(new File(pathname));
        }
        catch (BiffException e) {
            e.printStackTrace();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }

    public String readCell(int sheet,int row, int col)
    {
        Sheet s = workbook.getSheet(sheet);
        Cell c = s.getCell(row ,col);
        return c.getContents();
    }

    public void quitXLS()
    {
        if(workbook!=null){
            workbook.close();
        }
    }*/


    public void init() throws Exception {

        /*int i=0;
        System.out.println(i++);*/

        data.loadXLS("C:\\Users\\fchautem\\TestLab4techForm\\src\\test\\resources\\TestData.xls");
        //ws.setEncoding("UTF-8");
        driver.get(baseUrl);

        /*driver.findElement(By.id("view:_id1:_id46:_id57:TypeProvenance")).click();
        new Select(driver.findElement(By.id("view:_id1:_id46:_id57:TypeProvenance"))).selectByVisibleText("LACI (assurance-chômage)");*/
    }

    @Test(retryAnalyzer = RetryAnalyzer.class)
    public void testFormulaireAvecMesureLACI() throws Exception {
        init();
        donneesORPLACIAvecMesure(1);
        donneesPersonnelles(1);
        preferencesDeTravail(1);
        complementDInformation(1);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }

    @Test(retryAnalyzer = RetryAnalyzer.class)
    public void testFormulaireSansMesureLACI() throws Exception {
        init();
         donneesORPLACISansMesure(2);
        donneesPersonnelles(2);
        preferencesDeTravail(2);
        complementDInformation(2);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }

    @Test(retryAnalyzer = RetryAnalyzer.class)
    public void testFormulaireAvecMesureRI() throws Exception {
        init();
        donneesORPRIAvecMesure(3);
        donneesPersonnelles(3);
        preferencesDeTravail(3);
        complementDInformation(3);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }

    @Test(retryAnalyzer = RetryAnalyzer.class)
    public void testFormulaireSansMesureRI() throws Exception {
        init();
        donneesORPRISansMesure(4);
        donneesPersonnelles(4);
        preferencesDeTravail(4);
        complementDInformation(4);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }

    @Test(retryAnalyzer = RetryAnalyzer.class)
    public void testFormulaireAvecMesureAI() throws Exception {
        init();
        donneesORPAIAvecMesure(5);
        donneesPersonnelles(5);
        preferencesDeTravail(5);
        complementDInformation(5);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }

    @Test(retryAnalyzer = RetryAnalyzer.class)
    public void testFormulaireSansMesureAI() throws Exception {
        init();
        donneesORPAISansMesure(6);
        donneesPersonnelles(6);
        preferencesDeTravail(6);
        complementDInformation(6);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }

    @Test(retryAnalyzer = RetryAnalyzer.class)
    public void allMandatoryFields() throws Exception {
        init();
        donneesORPLACISansMesure(8);
        donneesPersonnelles(8);
        preferencesDeTravail(8);
        complementDInformation(8);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }

    @AfterClass
    public static void tearDown() throws Exception {

        driver.quit();
    }

    public void donneesORPLACIAvecMesure(int c) throws Exception {
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",data.readCell(0,0,c));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",data.readCell(0,1,c));
        selectKey(driver,"view:_id1:_id87:_id98:ORP",data.readCell(0,2,c));
        selectKey(driver,"view:_id1:first_caisse_refresh:_id109:CaisseChomage",data.readCell(0,3,c));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",data.readCell(0,4,c));
        insertKey(driver,"view:_id1:_id142:_id153:ORPDateInscription",data.readCell(0,5,c));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",data.readCell(0,6,c));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",data.readCell(0,7,c));
        insertKey(driver,"view:_id1:_id75:_id86:NumAssureORP",data.readCell(0,8,c));
        insertKey(driver,"view:_id1:_id157:_id168:NbreIndemnites",data.readCell(0,9,c));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",data.readCell(0,10,c));
        insertKey(driver,"view:_id1:stageCompany:_id222:MesureSuiviDetail",data.readCell(0,11,c));
        //System.out.println(readCell(0,11,c));
        //Thread.sleep(10000);
    }

    public void donneesORPLACISansMesure(int c) throws Exception {
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",data.readCell(0,0,c));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",data.readCell(0,1,c));
        selectKey(driver,"view:_id1:_id87:_id98:ORP",data.readCell(0,2,1));
        selectKey(driver,"view:_id1:first_caisse_refresh:_id109:CaisseChomage",data.readCell(0,3,c));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",data.readCell(0,4,c));
        insertKey(driver,"view:_id1:_id142:_id153:ORPDateInscription",data.readCell(0,5,c));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",data.readCell(0,6,c));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",data.readCell(0,7,c));
        insertKey(driver,"view:_id1:_id75:_id86:NumAssureORP",data.readCell(0,8,c));
        insertKey(driver,"view:_id1:_id157:_id168:NbreIndemnites",data.readCell(0,9,c));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",data.readCell(0,10,c));
    }


    public void donneesORPRIAvecMesure(int c) throws Exception {
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",data.readCell(0,0,c));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",data.readCell(0,1,c));
        selectKey(driver,"view:_id1:_id87:_id98:ORP",data.readCell(0,2,1));
        selectKey(driver,"view:_id1:first_caisse_refresh:_id109:CaisseChomage",data.readCell(0,3,c));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",data.readCell(0,4,c));
        insertKey(driver,"view:_id1:_id142:_id153:ORPDateInscription",data.readCell(0,5,c));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",data.readCell(0,6,c));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",data.readCell(0,7,c));
        insertKey(driver,"view:_id1:_id75:_id86:NumAssureORP",data.readCell(0,8,c));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",data.readCell(0,10,c));
        insertKey(driver,"view:_id1:stageCompany:_id222:MesureSuiviDetail",data.readCell(0,11,c));
    }

    public void donneesORPRISansMesure(int c) throws Exception {
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",data.readCell(0,0,c));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",data.readCell(0,1,c));
        selectKey(driver,"view:_id1:_id87:_id98:ORP",data.readCell(0,2,1));
        selectKey(driver,"view:_id1:first_caisse_refresh:_id109:CaisseChomage",data.readCell(0,3,c));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",data.readCell(0,4,c));
        insertKey(driver,"view:_id1:_id142:_id153:ORPDateInscription",data.readCell(0,5,c));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",data.readCell(0,6,c));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",data.readCell(0,7,c));
        insertKey(driver,"view:_id1:_id75:_id86:NumAssureORP",data.readCell(0,8,c));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",data.readCell(0,10,c));
    }


    public void donneesORPAIAvecMesure(int c) throws Exception {
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",data.readCell(0,0,c));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",data.readCell(0,1,c));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",data.readCell(0,4,c));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",data.readCell(0,6,c));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",data.readCell(0,7,c));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",data.readCell(0,10,c));
        insertKey(driver,"view:_id1:stageCompany:_id222:MesureSuiviDetail",data.readCell(0,11,c));
    }

    public void donneesORPAISansMesure(int c) throws Exception {
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",data.readCell(0,0,c));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",data.readCell(0,1,c));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",data.readCell(0,4,c));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",data.readCell(0,6,c));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",data.readCell(0,7,c));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",data.readCell(0,10,c));
    }

    public void donneesPersonnelles(int c) throws Exception {
        //Thread.sleep(30000);
        selectKey(driver,"view:_id1:_id353:_id364:CandidateTitre",data.readCell(0,12,c));
        insertKey(driver,"view:_id1:_id365:_id376:CandidatPrenom",data.readCell(0,13,c));
        insertKey(driver,"view:_id1:_id377:_id388:CandidatNom",data.readCell(0,14,c));
        insertKey(driver,"view:_id1:_id389:_id400:CandidatAdresse",data.readCell(0,15,c));
        insertKey(driver,"view:_id1:_id401:_id413:npaField",data.readCell(0,16,c));
        insertKey(driver,"view:_id1:candidatVilleField:_id426:_id427:_id428:cityField",data.readCell(0,17,c));
        insertKey(driver,"view:_id1:_id436:_id447:CandidatEmail",data.readCell(0,18,c));
        insertKey(driver,"view:_id1:_id449:_id460:CandidatPhone",data.readCell(0,19,c));
        insertKey(driver,"view:_id1:_id461:_id472:CandidatMobile",data.readCell(0,20,c));
        selectKey(driver,"view:_id1:_id473:_id484:CandidatEtatCivil",data.readCell(0,21,c));
        selectKey(driver,"view:_id1:_id485:_id496:CandidatNationalite",data.readCell(0,22,c));
        insertKey(driver,"view:_id1:_id497:_id508:CandidatAnniversaire",data.readCell(0,23,c));
        insertKey(driver,"view:_id1:_id511:_id522:CandidatWebsite",data.readCell(0,24,c));
        insertKey(driver,"view:_id1:_id523:_id534:CandidatNumAVS",data.readCell(0,25,c));
    }

    public void preferencesDeTravail(int c) throws Exception {
        //Thread.sleep(30000);
        selectKey(driver,"view:_id1:_id545:_id556:CandidatSecteurActivite",data.readCell(0,26,c));
        insertKey(driver,"view:_id1:_id557:_id568:CandidatJobActuel",data.readCell(0,27,c));
        insertKey(driver,"view:_id1:_id569:_id580:CandidatJobSouhait",data.readCell(0,28,c));
        selectKey(driver,"view:_id1:_id582:_id593:CandidatJobTauxOccup",data.readCell(0,29,c));
        selectKey(driver,"view:_id1:_id594:_id605:CandidatJobDispo",data.readCell(0,30,c));
    }

     public void complementDInformation(int c) throws Exception {
        insertKey(driver,"view:_id1:_id635:_id646:CandidatNbrePostulation3mois",data.readCell(0,31,c));
        insertKey(driver,"view:_id1:_id648:_id659:CandidatNbreRDVObtenus",data.readCell(0,32,c));
        insertKey(driver,"view:_id1:_id662:_id673:CandidatEntreprisesCibles",data.readCell(0,33,c));
        insertKey(driver,"view:_id1:companyRefresh:_id684:CandidatEntreprisesRencontrees",data.readCell(0,34,c));
        insertKey(driver,"view:_id1:_id685:_id696:IndisponibilitesPrevues",data.readCell(0,35,c));
        insertKey(driver,"view:_id1:_id697:_id708:Comments",data.readCell(0,36,c));
    }


    public void upload() throws Exception {
        Screen screen = new Screen();

        Thread.sleep(2000);
        screen.type(Key.END);
        screen.click(SikuliPath + "parcourir.png");
        Pattern p=new Pattern(SikuliPath + "path.png");
        p.targetOffset(-31,0);
        p.similar((float)0.8);
        screen.click(p);
        screen.paste(SikuliPath+"CV");
        screen.type(Key.ENTER);
        screen.click(SikuliPath + "CVTest.png");
        screen.click(SikuliPath + "open.png");
    }

    public void explicitWait(String id, String key)
    {
        (new WebDriverWait(driver,10)).until(ExpectedConditions.elementToBeClickable(By.id(id)));
    }



    public void send() throws Exception {
        boolean passed=false;
        while(!passed) {
            try {
                driver.findElement(By.id("view:_id1:btnSave")).click();
                passed = true;
            } catch (StaleElementReferenceException e) {
                System.out.println("StaleElementReferenceException");
            }
        }
    }

    private void insertKey(WebDriver driver, String id, String key) throws Exception
    {
        if(key=="") return;
        boolean passed=false;
        while(!passed) {
            try {
                driver.findElement(By.id(id)).sendKeys(key);
                passed = true;
            } catch (StaleElementReferenceException e) {
            }
        }
    }

    private void selectKey(WebDriver driver, String id, String key) throws Exception
    {
        if(key=="") return;
        boolean passed=false;
        while(!passed) {
            try {
                new Select(driver.findElement(By.id(id))).selectByVisibleText(key);
                passed = true;
            } catch (StaleElementReferenceException e) {
            }
        }
    }

    private boolean checkConfirmation() throws Exception
    {
        try {
            driver.findElement(By.xpath("//*[@id=\"view:_id1:UI_CONTENT_INSCRIPTION_CONFIRMATION_INTROTEXT1\"]"));
        } catch(NoSuchElementException e)
        {
            return false;
        }
        return true;
    }
}


