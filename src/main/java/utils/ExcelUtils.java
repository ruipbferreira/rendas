package utils;

import model.Fraction;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

public class ExcelUtils {

    public static List<Fraction> getFractionsFromFile(File exceFile, Properties props) throws Exception {
        List<Fraction> retObj = new ArrayList<Fraction>();
        FileInputStream stream = null;
        try {
            stream = new FileInputStream(exceFile);
            Workbook workbook = new HSSFWorkbook(stream);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            Integer startRow = Integer.parseInt(props.getProperty("startRow"));
            for (int i = 0;i < startRow;i++) {
                iterator.next();
            }
            while (iterator.hasNext()) {
                HSSFRow currentRow = (HSSFRow) iterator.next();
                HSSFCell cell = (HSSFCell) currentRow.getCell(10);
                if(cell != null) {
                    HSSFFont font = cell.getCellStyle().getFont(workbook);
                    short color = font.getColor();
                    if(color != 14 && !cell.getStringCellValue().startsWith("--")) {
                        Fraction fraction = new Fraction();
                        retObj.add(fraction);
                        //nome
                        fraction.setName(currentRow.getCell(Integer.parseInt(props.getProperty("nome"))).getStringCellValue());
                        //nif
                        fraction.setNif(currentRow.getCell(Integer.parseInt(props.getProperty("nif"))).getStringCellValue());
                        //valor
                        fraction.setValue(currentRow.getCell(Integer.parseInt(props.getProperty("valor"))).getNumericCellValue() + "€");
                        //morada
                        fraction.setAddress(currentRow.getCell(Integer.parseInt(props.getProperty("morada"))).getStringCellValue());
                        //codigofreguesia
                        fraction.setAddressCode(currentRow.getCell(Integer.parseInt(props.getProperty("codigofreguesia"))).getStringCellValue());
                        //codigoartigo
                        double value = currentRow.getCell(Integer.parseInt(props.getProperty("codigoartigo"))).getNumericCellValue();
                        fraction.setArticleCode((int)value + "");
                        //codigofracao
                        fraction.setFractionCode(currentRow.getCell(Integer.parseInt(props.getProperty("codigofracao"))).getStringCellValue());
                        //quotaparte
                        fraction.setShare(currentRow.getCell(Integer.parseInt(props.getProperty("quotaparte"))).getStringCellValue());

                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new Exception("Erro ao ler ficheiro Excel");
        } finally {
            try {
                stream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return retObj;
    }
}
