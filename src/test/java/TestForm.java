import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import jdk.nashorn.internal.ir.annotations.Ignore;

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

//import java.nio.charset.Charset;*/

public class TestForm {

    private static WebDriver driver;
    private static String Chromedriverpath;
    private static String Chromedriverlog;
    private static String SikuliPath;
    private static String baseUrl;
    private static int timeout=5000;
    private static final String MESSAGE_PROVIDER = "MessageProvider";
    private static DataManagement dm=new DataManagement();
    private static int index = 0 ;
    private static DataForm[] data;

    private static final int rowMin=1;
    private static final int rowMax=6;
    private static final int colMin=0;
    private static final int colMax=36;

    //private static Workbook workbook;

    //private static WorkbookSettings ws = new WorkbookSettings();

    //private DataManagement data=new DataManagement();
    private MessageProvider mp = new MessageProvider();

    public static void initData()
    {
        dm.loadXLS("C:\\Users\\fchautem\\TestLab4techForm\\src\\test\\resources\\TestData.xls");
        //data=new DataForm[2];
        data=dm.readAll(rowMin,rowMax,colMin,colMax);
        dm.quitXLS();
        //System.out.println("LONGUEUR : " + data.length);
        for(int i=0;i<rowMax-rowMin+1;i++)data[i].print();
    }

    public static class IteratorData implements Iterator<Object>
    {
        private IteratorData()
        {
            initData();
        }

        public boolean hasNext() {
            return index < data.length;
        }

        public DataForm next() {
            System.out.println("NEXT : " + index);
            DataForm df = data[index];
            index++ ;
            return df;
        }

        public void remove() {
            throw new UnsupportedOperationException() ;
        }
    }

    @DataProvider(name=MESSAGE_PROVIDER)
    public static Iterator<Object> provideMessage() {
        return new IteratorData();
    }


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

    public void init() throws Exception {

        /*int i=0;
        System.out.println(i++);*/

        //data.loadXLS("C:\\Users\\fchautem\\TestLab4techForm\\src\\test\\resources\\TestData.xls");
        //ws.setEncoding("UTF-8");
        driver.get(baseUrl);
        //mp.init();
        /*driver.findElement(By.id("view:_id1:_id46:_id57:TypeProvenance")).click();
        new Select(driver.findElement(By.id("view:_id1:_id46:_id57:TypeProvenance"))).selectByVisibleText("LACI (assurance-chômage)");*/
    }

    /*@Ignore(retryAnalyzer = RetryAnalyzer.class)
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

    @Ignore(retryAnalyzer = RetryAnalyzer.class)
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

    @Ignore(retryAnalyzer = RetryAnalyzer.class)
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

    @Ignore(retryAnalyzer = RetryAnalyzer.class)
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

    @Ignore(retryAnalyzer = RetryAnalyzer.class)
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

    @Ignore(retryAnalyzer = RetryAnalyzer.class)
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
    }*/

    @Test(retryAnalyzer = RetryAnalyzer.class,dataProvider="MessageProvider")
    public void allMandatoryFields(DataForm df) throws Exception {
        init();
        donneesORP(df);
        donneesPersonnelles(df);
        preferencesDeTravail(df);
        complementDInformation(df);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
    }

    /*@Test()
    public void allMandatoryFields() throws Exception {
        System.out.println("INIT");
        init();
        DataForm df;
        df=data.readAll(0,1,0,36)[0];

        testDataSet(df); System.out.println("testDataSet");

        /*donneesORPLACISansMesure(8);
        donneesPersonnelles(8);
        preferencesDeTravail(8);
        complementDInformation(8);
        upload();
        send();
        Assert.assertTrue(checkConfirmation());
        data.quitXLS();
    }*/

    @AfterClass
    public static void tearDown() throws Exception {

        driver.quit();
    }

    public void donneesORP(DataForm df) throws Exception {
        System.out.println("Execution Numero : " + index);
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",df.get("TypeProvenance"));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",df.get("Canton"));
        selectKey(driver,"view:_id1:_id87:_id98:ORP",df.get("ORP"));
        selectKey(driver,"view:_id1:first_caisse_refresh:_id109:CaisseChomage",df.get("CaisseChomage"));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",df.get("TypePrestation"));
        insertKey(driver,"view:_id1:_id142:_id153:ORPDateInscription",df.get("ORPDateInscription"));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",df.get("CPEmail"));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",df.get("MesureLab4tech"));
        insertKey(driver,"view:_id1:_id75:_id86:NumAssureORP",df.get("NumAssureORP"));
        insertKey(driver,"view:_id1:_id157:_id168:NbreIndemnites",df.get("NbreIndemnites"));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",df.get("MesureSuiviReponse"));
        insertKey(driver,"view:_id1:stageCompany:_id222:MesureSuiviDetail",df.get("MesureSuiviDetail"));
    }


    public void donneesPersonnelles(DataForm df) throws Exception {
        //Thread.sleep(30000);
        selectKey(driver,"view:_id1:_id353:_id364:CandidateTitre",df.get("CandidateTitre"));
        insertKey(driver,"view:_id1:_id365:_id376:CandidatPrenom",df.get("CandidatPrenom"));
        insertKey(driver,"view:_id1:_id377:_id388:CandidatNom",df.get("CandidatNom"));
        insertKey(driver,"view:_id1:_id389:_id400:CandidatAdresse",df.get("CandidatAdresse"));
        insertKey(driver,"view:_id1:_id401:_id413:npaField",df.get("npaField"));
        insertKey(driver,"view:_id1:candidatVilleField:_id426:_id427:_id428:cityField",df.get("candidatVilleField"));
        insertKey(driver,"view:_id1:_id436:_id447:CandidatEmail",df.get("CandidatEmail"));
        insertKey(driver,"view:_id1:_id449:_id460:CandidatPhone",df.get("CandidatPhone"));
        insertKey(driver,"view:_id1:_id461:_id472:CandidatMobile",df.get("CandidatMobile"));
        selectKey(driver,"view:_id1:_id473:_id484:CandidatEtatCivil",df.get("CandidatEtatCivil"));
        selectKey(driver,"view:_id1:_id485:_id496:CandidatNationalite",df.get("CandidatNationalite"));
        insertKey(driver,"view:_id1:_id497:_id508:CandidatAnniversaire",df.get("CandidatAnniversaire"));
        insertKey(driver,"view:_id1:_id511:_id522:CandidatWebsite",df.get("CandidatWebsite"));
        insertKey(driver,"view:_id1:_id523:_id534:CandidatNumAVS",df.get("CandidatNumAVS"));
    }



    /*public void donneesORPLACIAvecMesure(DataForm df) throws Exception {
        selectKey(driver,"view:_id1:_id46:_id57:TypeProvenance",df.get("TypeProvenance"));
        selectKey(driver,"view:_id1:_id61:_id72:Canton",df.get("Canton"));
        selectKey(driver,"view:_id1:_id87:_id98:ORP",df.get("ORP"));
        selectKey(driver,"view:_id1:first_caisse_refresh:_id109:CaisseChomage",df.get("CaisseChomage"));
        selectKey(driver,"view:_id1:_id114:_id125:TypePrestation",df.get("TypePrestation"));
        insertKey(driver,"view:_id1:_id142:_id153:ORPDateInscription",df.get("ORPDateInscription"));
        insertKey(driver,"view:_id1:_id170:_id181:CPEmail",df.get("CPEmail"));
        selectKey(driver,"view:_id1:_id185:_id196:MesureLab4tech",df.get("MesureLab4tech"));
        insertKey(driver,"view:_id1:_id75:_id86:NumAssureORP",df.get("NumAssureORP"));
        insertKey(driver,"view:_id1:_id157:_id168:NbreIndemnites",df.get("NbreIndemnites"));
        selectKey(driver,"view:_id1:_id198:_id209:MesureSuiviReponse",df.get("MesureSuiviReponse"));
        insertKey(driver,"view:_id1:stageCompany:_id222:MesureSuiviDetail",df.get("MesureSuiviDetail"));
        //System.out.println(readCell(0,11,c));
        //Thread.sleep(10000);
    }

    public void donneesORPLACISansMesure(DataForm df) throws Exception {
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
    }*/

    public void preferencesDeTravail(DataForm df) throws Exception {
        //Thread.sleep(30000);
        selectKey(driver,"view:_id1:_id545:_id556:CandidatSecteurActivite",df.get("CandidatSecteurActivite"));
        insertKey(driver,"view:_id1:_id557:_id568:CandidatJobActuel",df.get("CandidatJobActuel"));
        insertKey(driver,"view:_id1:_id569:_id580:CandidatJobSouhait",df.get("CandidatJobSouhait"));
        selectKey(driver,"view:_id1:_id582:_id593:CandidatJobTauxOccup",df.get("CandidatJobTauxOccup"));
        selectKey(driver,"view:_id1:_id594:_id605:CandidatJobDispo",df.get("CandidatJobDispo"));
    }

     public void complementDInformation(DataForm df) throws Exception {
        try {
            insertKey(driver, "view:_id1:_id635:_id646:CandidatNbrePostulation3mois", df.get("CandidatNbrePostulation3mois"));
            System.out.println("CandidatNbrePostulation3mois");
            System.out.println(df.get("CandidatNbrePostulation3mois"));
            insertKey(driver, "view:_id1:_id648:_id659:CandidatNbreRDVObtenus", df.get("CandidatNbreRDVObtenus"));
            System.out.println("CandidatNbreRDVObtenus");
            System.out.println(df.get("CandidatNbreRDVObtenus"));
            insertKey(driver, "view:_id1:_id662:_id673:CandidatEntreprisesCibles", df.get("CandidatEntreprisesCibles"));
            System.out.println("CandidatEntreprisesCibles");
            System.out.println(df.get("CandidatEntreprisesCibles"));
            insertKey(driver, "view:_id1:companyRefresh:_id684:CandidatEntreprisesRencontrees", df.get("CandidatEntreprisesRencontrees"));
            System.out.println("CandidatEntreprisesRencontrees");
            System.out.println(df.get("CandidatEntreprisesRencontrees"));
            insertKey(driver, "view:_id1:_id685:_id696:IndisponibilitesPrevues", df.get("IndisponibilitesPrevues"));
            System.out.println("IndisponibilitesPrevues");
            System.out.println(df.get("IndisponibilitesPrevues"));
            insertKey(driver, "view:_id1:_id697:_id708:Comments", df.get("Comments"));
            System.out.println("Comments");
            System.out.println(df.get("Comments"));
        }catch(Exception e)
        {
            System.out.println("MERDE UNE ERREUR : " + e.getStackTrace());
        }
    }

    public void upload() throws Exception {
        System.out.println("Je suis dans la méthode upload");
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


