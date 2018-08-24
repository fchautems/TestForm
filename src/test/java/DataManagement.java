import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;

public class DataManagement {

    private Workbook workbook;

    private WorkbookSettings ws = new WorkbookSettings();

    public static final String DATA_PROVIDER = "DataProvider";

    public void loadXLS(String pathname) {

        workbook = null;
        ws.setEncoding("UTF-8");
        try {
            /* Récupération du classeur Excel (en lecture) */
            workbook = Workbook.getWorkbook(new File(pathname));
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public DataForm[] readAll(int rowMin, int rowMax, int colMin, int colMax) {
        DataForm[] data = new DataForm[rowMax - rowMin + 1];
        for (int i = rowMin; i <= rowMax; i++) {
            //System.out.println("I MOINS ROW MIN : "+ (i-rowMin));
            data[i - rowMin] = new DataForm();
            for (int j = colMin; j <= colMax; j++) {
                data[i - rowMin].add(readCell(0, 0, j), readCell(0, i, j));
            }
        }
        return data;
    }

    public String readCell(int sheet, int row, int col) {
        /* Un fichier excel est composé de plusieurs feuilles, on y accède de la manière suivante*/
        Sheet s = workbook.getSheet(sheet);
        /* On accède aux cellules avec la méthode getCell(indiceColonne, indiceLigne) */
        Cell c = s.getCell(col, row);
        return c.getContents();
    }

    public void quitXLS() {
        if (workbook != null) {
            workbook.close();
        }
    }

}
