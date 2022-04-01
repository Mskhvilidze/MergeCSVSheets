import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class Main {
    /*static String dirPath = "C:\\Users\\p05865\\Desktop\\JavaTest\\";
    static String path = "P:/APIForOle/javaMergeCSV/DealersList.xlsx";*/

    public static void main(String[] args) throws IOException, URISyntaxException {
        mergeSheets();
    }

    static void mergeSheets() throws IOException, URISyntaxException {
        List<Sheet> list;
        WorkBookClass workBookClass = new WorkBookClass();
        //jar can find itself
        URL url = workBookClass.getLocation(Main.class);
        workBookClass.urlToFile(url.toURI().toString());
        String location = url.toString().replace("MergeCSVFils.jar", "");
        String dirPath = location.replace("file:/", "") + "test/";
        String path = location.replace("file:/", "") + "javaMergeCSV/DealersList.xlsx";
        //Name
        final String fileName = "test.xlsx";
        //the path where the file is placed
        String path1 = location.replace("file:/", "") + "test/" + workBookClass.createFile(dirPath, fileName);

        FileOutputStream outputStream = new FileOutputStream(path1);

        Iterator<Sheet> sheets = workBookClass.inputWorkbook(path);
        list = workBookClass.getList(sheets);

        Workbook workbookOutput = workBookClass.outputWorkbook();
        XSSFSheet outputSheet = (XSSFSheet) workbookOutput.createSheet("Test");

        workBookClass.setColumnWidth(outputSheet);
        //copy Sheets
        copySheets(workBookClass, outputSheet, list);
        workbookOutput.write(outputStream);
    }

    private static void copySheets(WorkBookClass workBookClass, Sheet outputSheet, List<Sheet> list) {
        //Create Cell and Row
        XSSFCell cell = null;
        XSSFRow row = null;
        List<Integer> currentList = new ArrayList<>();
        int currentLine = 1;
        int getColumnIndex = getColumnIndex(list);
        //
        for (int i = 0; i < list.size(); i++) {
            int temp = 1;
            currentLine--;
            // First line is ignored
            if (i >= 1) {
                temp = 2;
            }
            //Row index calculate
            int getRowIndex = getRowIndex(list, i);

            for (int k = temp; k <= getRowIndex; k++, currentLine++) {
                row = (XSSFRow) outputSheet.createRow(currentLine);
                for (int j = 0; j < getColumnIndex; j++) {
                    cell = row.createCell(j);
                    if (list.get(i).getRow(k) == null) {
                        cell.setCellValue("");
                    } else {
                        cell.setCellValue(workBookClass.getDataFormatter().formatCellValue(list.get(i).getRow(k).getCell(j)));
                        if (workBookClass.getDataFormatter().formatCellValue(list.get(i).getRow(k).getCell(j)).equals("FR57669200784")) {
                            currentList.add(k);
                        }
                    }
                }
            }
        }
        setFlag(currentList, outputSheet, getColumnIndex);
    }

    private static void setFlag(List<Integer> currentList, Sheet outputSheet, int getColumnIndex) {
        for (int i = 0; i < currentList.size(); i++) {
            outputSheet.getRow(currentList.get(i)).getCell(getColumnIndex - 1).setCellValue("ja");
        }
    }

    /**
     * Retun number of Cell
     *
     * @param list
     * @return
     */
    private static int getColumnIndex(List<Sheet> list) {
        if (list.isEmpty()) {
            throw new IllegalArgumentException("List is empty! ");
        }
        return list.get(0).getRow(1).getPhysicalNumberOfCells();
    }

    /**
     * return number of Row
     *
     * @param list
     * @param count
     * @return
     */
    private static int getRowIndex(List<Sheet> list, int count) {
        if (list.isEmpty()) {
            throw new IllegalArgumentException("List is empty! ");
        }

        if (count == list.size() - 1) {
            return list.get(count).getPhysicalNumberOfRows() - 2;
        } else
            return list.get(count).getPhysicalNumberOfRows() - 1;
    }
}