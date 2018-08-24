import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.sikuli.script.Key;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.Iterator;
import java.util.concurrent.TimeUnit;

public class TestForm {

    private static WebDriver driver;
    private static String Chromedriverpath;
    private static String Chromedriverlog;
    private static String SikuliPath;
    private static String baseUrl;
    private static int timeout = 5000;
    private static final String MESSAGE_PROVIDER = "MessageProvider";
    private static DataManagement dm = new DataManagement();
    private static int index = 0;
    private static DataForm[] data;
    private static final int rowMin = 1;
    private static final int rowMax = 6;
    private static final int colMin = 0;
    private static final int colMax = 36;

    //private MessageProvider mp = new MessageProvider();

    public static void initData() {
        dm.loadXLS("C:\\Users\\fchautem\\TestLab4techForm\\src\\test\\resources\\TestData.xls");
        data = dm.readAll(rowMin, rowMax, colMin, colMax);
        dm.quitXLS();
        for (int i = 0; i < rowMax - rowMin + 1; i++) data[i].print();
    }

    public static class IteratorData implements Iterator<Object> {
        private IteratorData() {
            initData();
        }

        public boolean hasNext() {
            return index < data.length;
        }

        public DataForm next() {
            DataForm df = data[index];
            index++;
            return df;
        }

        public void remove() {
            throw new UnsupportedOperationException();
        }
    }

    @DataProvider(name = MESSAGE_PROVIDER)
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
        driver.manage().timeouts().implicitlyWait(timeout, TimeUnit.MILLISECONDS);
        //java.nio.charset.StandardCharsets.UTF_8.encode(string);
    }

    public void init() throws Exception {
        //ws.setEncoding("UTF-8");
        driver.get(baseUrl);
    }

    @Test(dataProvider = "MessageProvider") //retryAnalyzer = RetryAnalyzer.class,
    public void allMandatoryFields(DataForm df) throws Exception {
        init();
        upload(df);
        donneesORP(df);
        donneesPersonnelles(df);
        preferencesDeTravail(df);
        complementDInformation(df);
        send();
        Assert.assertTrue(checkConfirmation());
    }

    @AfterClass
    public static void tearDown() throws Exception {

        driver.quit();
    }

    public void donneesORP(DataForm df) throws Exception {
        selectKey(driver, "view:_id1:_id46:_id57:TypeProvenance", df.get("TypeProvenance"));
        selectKey(driver, "view:_id1:_id61:_id72:Canton", df.get("Canton"));
        selectKey(driver, "view:_id1:_id87:_id98:ORP", df.get("ORP"));
        selectKey(driver, "view:_id1:first_caisse_refresh:_id109:CaisseChomage", df.get("CaisseChomage"));
        selectKey(driver, "view:_id1:_id114:_id125:TypePrestation", df.get("TypePrestation"));
        insertKey(driver, "view:_id1:_id142:_id153:ORPDateInscription", df.get("ORPDateInscription"));
        insertKey(driver, "view:_id1:_id170:_id181:CPEmail", df.get("CPEmail"));
        selectKey(driver, "view:_id1:_id185:_id196:MesureLab4tech", df.get("MesureLab4tech"));
        insertKey(driver, "view:_id1:_id75:_id86:NumAssureORP", df.get("NumAssureORP"));
        insertKey(driver, "view:_id1:_id157:_id168:NbreIndemnites", df.get("NbreIndemnites"));
        selectKey(driver, "view:_id1:_id198:_id209:MesureSuiviReponse", df.get("MesureSuiviReponse"));
        insertKey(driver, "view:_id1:stageCompany:_id222:MesureSuiviDetail", df.get("MesureSuiviDetail"));
    }

    public void donneesPersonnelles(DataForm df) throws Exception {
        selectKey(driver, "view:_id1:_id353:_id364:CandidateTitre", df.get("CandidateTitre"));
        insertKey(driver, "view:_id1:_id365:_id376:CandidatPrenom", df.get("CandidatPrenom"));
        insertKey(driver, "view:_id1:_id377:_id388:CandidatNom", df.get("CandidatNom"));
        insertKey(driver, "view:_id1:_id389:_id400:CandidatAdresse", df.get("CandidatAdresse"));
        insertKey(driver, "view:_id1:_id401:_id413:npaField", df.get("npaField"));
        insertKey(driver, "view:_id1:candidatVilleField:_id426:_id427:_id428:cityField", df.get("candidatVilleField"));
        insertKey(driver, "view:_id1:_id436:_id447:CandidatEmail", df.get("CandidatEmail"));
        insertKey(driver, "view:_id1:_id449:_id460:CandidatPhone", df.get("CandidatPhone"));
        insertKey(driver, "view:_id1:_id461:_id472:CandidatMobile", df.get("CandidatMobile"));
        selectKey(driver, "view:_id1:_id473:_id484:CandidatEtatCivil", df.get("CandidatEtatCivil"));
        selectKey(driver, "view:_id1:_id485:_id496:CandidatNationalite", df.get("CandidatNationalite"));
        insertKey(driver, "view:_id1:_id497:_id508:CandidatAnniversaire", df.get("CandidatAnniversaire"));
        insertKey(driver, "view:_id1:_id511:_id522:CandidatWebsite", df.get("CandidatWebsite"));
        insertKey(driver, "view:_id1:_id523:_id534:CandidatNumAVS", df.get("CandidatNumAVS"));
    }

    public void preferencesDeTravail(DataForm df) throws Exception {
        selectKey(driver, "view:_id1:_id545:_id556:CandidatSecteurActivite", df.get("CandidatSecteurActivite"));
        insertKey(driver, "view:_id1:_id557:_id568:CandidatJobActuel", df.get("CandidatJobActuel"));
        insertKey(driver, "view:_id1:_id569:_id580:CandidatJobSouhait", df.get("CandidatJobSouhait"));
        selectKey(driver, "view:_id1:_id582:_id593:CandidatJobTauxOccup", df.get("CandidatJobTauxOccup"));
        selectKey(driver, "view:_id1:_id594:_id605:CandidatJobDispo", df.get("CandidatJobDispo"));
    }

    public void complementDInformation(DataForm df) throws Exception {
        insertKey(driver, "view:_id1:_id635:_id646:CandidatNbrePostulation3mois", df.get("CandidatNbrePostulation3mois"));
        insertKey(driver, "view:_id1:_id648:_id659:CandidatNbreRDVObtenus", df.get("CandidatNbreRDVObtenus"));
        insertKey(driver, "view:_id1:_id662:_id673:CandidatEntreprisesCibles", df.get("CandidatEntreprisesCibles"));
        insertKey(driver, "view:_id1:companyRefresh:_id684:CandidatEntreprisesRencontrees", df.get("CandidatEntreprisesRencontrees"));
        insertKey(driver, "view:_id1:_id685:_id696:IndisponibilitesPrevues", df.get("IndisponibilitesPrevues"));
        insertKey(driver, "view:_id1:_id697:_id708:Comments", df.get("Comments"));
    }

    public void upload(DataForm df) throws Exception {

        boolean passed = false;
        Screen screen = new Screen();

        while (!passed) {
            try {
                screen.click(SikuliPath + "inscription.png");
                screen.type(Key.END);
                screen.click(SikuliPath + "parcourir.png");
                Pattern p = new Pattern(SikuliPath + "path.png");
                p.targetOffset(-31, 0);
                p.similar((float) 0.8);
                screen.click(p);
                screen.paste(SikuliPath + "CV");
                screen.type(Key.ENTER);
                screen.click(SikuliPath + "CVTest.png");
                screen.click(SikuliPath + "open.png");
                selectKey(driver, "view:_id1:_id46:_id57:TypeProvenance", df.get("TypeProvenance"));
                passed = true;
            } catch (Exception e) {
                screen.keyDown(Key.F5);
                Thread.sleep(500);
                screen.keyUp(Key.F5);
            }
        }
    }

    public void send() throws Exception {
        boolean passed = false;
        while (!passed) {
            try {
                driver.findElement(By.id("view:_id1:btnSave")).click();
                passed = true;
            } catch (StaleElementReferenceException e) {
            }
        }
    }

    private void insertKey(WebDriver driver, String id, String key) throws Exception {
        if (key == "") return;
        boolean passed = false;
        while (!passed) {
            try {
                driver.findElement(By.id(id)).sendKeys(key);
                passed = true;
            } catch (StaleElementReferenceException e) {
            }
        }
    }

    private void selectKey(WebDriver driver, String id, String key) throws Exception {
        if (key == "") return;
        boolean passed = false;
        while (!passed) {
            try {
                new Select(driver.findElement(By.id(id))).selectByVisibleText(key);
                passed = true;
            } catch (StaleElementReferenceException e) {
            }
        }
    }

    private boolean checkConfirmation() throws Exception {
        try {
            driver.findElement(By.xpath("//*[@id=\"view:_id1:UI_CONTENT_INSCRIPTION_CONFIRMATION_INTROTEXT1\"]"));
        } catch (NoSuchElementException e) {
            return false;
        }
        return true;
    }
}