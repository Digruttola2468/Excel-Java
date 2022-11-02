import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Excel {


    private static final String nombreArchivo = "Java.xlsx";

    public static void main(String[] args) {
        Workbook libro = new XSSFWorkbook();
        Sheet hoja = libro.createSheet("Hoja Java");

        Row primeraFila = hoja.createRow(0);

        Cell primeraCelda = primeraFila.createCell(0);
        primeraCelda.setCellValue("Primera");
        Cell segundaCelda = primeraFila.createCell(1);
        segundaCelda.setCellValue("Segunda");
        Cell terceraCelda = primeraFila.createCell(2);
        terceraCelda.setCellValue("Tercera");
        Cell cuartaCelda = primeraFila.createCell(3);
        cuartaCelda.setCellValue("Cuarta");

        Row segundaFila = hoja.createRow(1);

        Cell quintaCelda = segundaFila.createCell(0);
        quintaCelda.setCellValue("Quinta");
        Cell sextaCelda = segundaFila.createCell(1);
        sextaCelda.setCellValue("Sexta");
        Cell septimaCelda = segundaFila.createCell(2);
        septimaCelda.setCellValue("Septima");
        Cell octabaCelda = segundaFila.createCell(3);
        octabaCelda.setCellValue("Octaba");

        File directorioActual = new File(".");
        String ubicacion = directorioActual.getAbsolutePath();
        String ubicacionArchivoSalida = ubicacion.substring(0, ubicacion.length() - 1) + nombreArchivo;
        FileOutputStream outputStream;
        try {
            outputStream = new FileOutputStream(ubicacionArchivoSalida);
            libro.write(outputStream);
            libro.close();
            System.out.println("Libro guardado correctamente");
        } catch (FileNotFoundException ex) {
            System.out.println("Error de filenotfound");
        } catch (IOException ex) {
            System.out.println("Error de IOException");
        }
    }
}
