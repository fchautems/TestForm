import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

public class DataManagement {

    private Workbook workbook;

    private List<List<String>> list;

    private Iterator<List<String>> iterator;

    private WorkbookSettings ws = new WorkbookSettings();

    public void loadXLS(String pathname) {

        workbook=null;
        ws.setEncoding("UTF-8");
        try {
            /* Récupération du classeur Excel (en lecture) */
            workbook = Workbook.getWorkbook(new File(pathname));
        }
        catch (BiffException e) {
            e.printStackTrace();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        /*finally {
            if(workbook!=null){
                workbook.close();
            }
        }*/
    }

    public Iterator<List<String>> readAll(int rowMin, int rowMax, int colMin, int colMax)
    {
        List<String> ligne = null;
        for(int i=rowMin;i<rowMax;i++)
        {
            for(int j=colMin;j<colMax;j++)
            {
                ligne.add(readCell(0,i,j));
            }
            list.add(ligne);
        }
        return list.iterator();
    }

    public String readCell(int sheet,int row, int col)
    {
        /* Un fichier excel est composé de plusieurs feuilles, on y accède de la manière suivante*/
        Sheet s = workbook.getSheet(sheet);
        /* On accède aux cellules avec la méthode getCell(indiceColonne, indiceLigne) */
        Cell c = s.getCell(row ,col);
        return c.getContents();
    }

    public void quitXLS()
    {
        if(workbook!=null){
            workbook.close();
        }
    }

}
