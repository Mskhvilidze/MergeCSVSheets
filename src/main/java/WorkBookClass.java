
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class WorkBookClass {
    private List<Sheet> list;

    public Iterator<Sheet> inputWorkbook(String path) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(path));
        Workbook inputWorkbook = new XSSFWorkbook(inputStream);
        Iterator<Sheet> iterator = inputWorkbook.iterator();
        return iterator;
    }

    private Sheet inputWorkbookSheet(Workbook workbook) {
        Sheet sheetDepth = workbook.getSheetAt(0);
        return sheetDepth;
    }

    public Workbook outputWorkbook() throws FileNotFoundException {
        Workbook outputWorkbook = new XSSFWorkbook();
        return outputWorkbook;
    }

    public void setColumnWidth(Sheet sheet) {
        int[] arr = sheet.getColumnBreaks();
        sheet.setColumnWidth(0, 8500);
    }

    public CellStyle getStyleFontBold(CellStyle styleFontBold) {
        return styleFontBold;
    }

    public CellStyle getStyleCenter(Workbook workbook) {
        CellStyle cellStyleCenter = workbook.createCellStyle();
        cellStyleCenter.setAlignment(HorizontalAlignment.CENTER);
        return cellStyleCenter;
    }

    public CellStyle getStyleColor(Workbook workbook) {
        CellStyle cellStyleColor = workbook.createCellStyle();
        cellStyleColor.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cellStyleColor.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyleColor;
    }

    public CellStyle getFontBold(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        return cellStyle;
    }

    public CellStyle getFontBoldAndUnderiline(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setUnderline(Font.U_DOUBLE);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        return cellStyle;
    }

    public String createFile(String dirtPath, String fileName) {
        Path filePath = Paths.get(dirtPath + fileName);
        File file = new File(fileName);
        if (!file.exists()) {
            try {
                Files.createFile(filePath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return file.getName();
    }

    public DataFormatter getDataFormatter(){
        DataFormatter dataFormatter = new DataFormatter();
        return dataFormatter;
    }

    public List<Sheet> getList(Iterator<Sheet> iterator){
        list = new ArrayList<>();
        while(iterator.hasNext()){
            list.add(iterator.next());
        }
        return list;
    }

    /**
     * Gets the base location of the given class.
     * <p>
     * If the class is directly on the file system (e.g.,
     * "/path/to/my/package/MyClass.class") then it will return the base directory
     * (e.g., "file:/path/to").
     * </p>
     * <p>
     * If the class is within a JAR file (e.g.,
     * "/path/to/my-jar.jar!/my/package/MyClass.class") then it will return the
     * path to the JAR (e.g., "file:/path/to/my-jar.jar").
     * </p>
     *
     * @param c The class whose location is desired.
     */
    public  URL getLocation(final Class<?> c)  {
        if (c == null) return null; // could not load the class
        // try the easy way first
        try {
            final URL codeSourceLocation =
                    c.getProtectionDomain().getCodeSource().getLocation();
            if (codeSourceLocation != null) return codeSourceLocation;
        }
        catch (final SecurityException e) {
            // NB: Cannot access protection domain.
        }
        catch (final NullPointerException e) {
            // NB: Protection domain or code source is null.
        }

        // NB: The easy way failed, so we try the hard way. We ask for the class
        // itself as a resource, then strip the class's path from the URL string,
        // leaving the base path.

        // get the class's raw resource path
        final URL classResource = c.getResource(c.getSimpleName() + ".class");
        if (classResource == null) return null; // cannot find class resource

        final String url = classResource.toString();
        final String suffix = c.getCanonicalName().replace('.', '/') + ".class";
        if (!url.endsWith(suffix)) return null; // weird URL

        // strip the class's path from the URL string
        final String base = url.substring(0, url.length() - suffix.length());

        String path = base;

        // remove the "jar:" prefix and "!/" suffix, if present
        if (path.startsWith("jar:")) path = path.substring(4, path.length() - 2);

        try {
            return new URL(path);
        }
        catch (final MalformedURLException e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * Converts the given {@link URL} to its corresponding {@link File}.
     * <p>
     * This method is similar to calling {@code new File(url.toURI())} except that
     * it also handles "jar:file:" URLs, returning the path to the JAR file.
     * </p>
     *
     * @param url The URL to convert.
     * @return A file path suitable for use with e.g. {@link FileInputStream}
     * @throws IllegalArgumentException if the URL does not correspond to a file.
     */
    public  File urlToFile(final URL url) {
        return url == null ? null : urlToFile(url.toString());
    }

    /**
     * Converts the given URL string to its corresponding {@link File}.
     *
     * @param url The URL to convert.
     * @return A file path suitable for use with e.g. {@link FileInputStream}
     * @throws IllegalArgumentException if the URL does not correspond to a file.
     */
    public  File urlToFile(final String url) {
        String path = url;
        if (path.startsWith("jar:")) {
            // remove "jar:" prefix and "!/" suffix
            final int index = path.indexOf("!/");
            path = path.substring(4, index);
        }
        try {
            if (path.matches("file:[A-Za-z]:.*")) {
                path = "file:/" + path.substring(5);
            }
            return new File(new URL(path).toURI());
        }
        catch (final MalformedURLException e) {
            // NB: URL is not completely well-formed.
        }
        catch (final URISyntaxException e) {
            // NB: URL is not completely well-formed.
        }
        if (path.startsWith("file:")) {
            // pass through the URL as-is, minus "file:" prefix
            path = path.substring(5);
            return new File(path);
        }
        throw new IllegalArgumentException("Invalid URL: " + url);
    }
}
