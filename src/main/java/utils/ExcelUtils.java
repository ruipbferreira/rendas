package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import model.Fraction;

public class ExcelUtils {

    public static List<Fraction> getFractionsFromFile(File exceFile, Map<String, String> props) throws Exception {
        List<Fraction> retObj = new ArrayList<Fraction>();
        FileInputStream stream = null;
        try {
            stream = new FileInputStream(exceFile);
            Workbook workbook = new HSSFWorkbook(stream);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            Integer startRow = Integer.parseInt(props.get("startRow"));
            for (int i = 0;i < startRow;i++) {
                iterator.next();
            }
            while (iterator.hasNext()) {
                HSSFRow currentRow = (HSSFRow) iterator.next();
                HSSFCell cell = (HSSFCell) currentRow.getCell(10);
                if(cell != null) {
                    String isToGenerate = currentRow.getCell(Integer.parseInt(props.get("isToGenerate"))).getStringCellValue();
                    if("S".equalsIgnoreCase(isToGenerate)) {
                        Fraction fraction = new Fraction();
                        retObj.add(fraction);
                        //nome
                        fraction.setName(currentRow.getCell(Integer.parseInt(props.get("nome"))).getStringCellValue());
                        //nif
                        HSSFCell niffCell = currentRow.getCell(Integer.parseInt(props.get("nif")));
                        fraction.setNif(niffCell.getCellType() == 1 ? niffCell.getStringCellValue() : (int)niffCell.getNumericCellValue() + "");
                        //valor
                        fraction.setValue(String.format("%.2f", currentRow.getCell(Integer.parseInt(props.get("valor"))).getNumericCellValue()) + "€");
                        //morada
                        fraction.setAddress(currentRow.getCell(Integer.parseInt(props.get("morada"))).getStringCellValue());
                        //codigofreguesia
                        fraction.setAddressCode(currentRow.getCell(Integer.parseInt(props.get("codigofreguesia"))).getStringCellValue());
                        //codigoartigo
                        double value = currentRow.getCell(Integer.parseInt(props.get("codigoartigo"))).getNumericCellValue();
                        fraction.setArticleCode((int)value + "");
                        //codigofracao
                        fraction.setFractionCode(currentRow.getCell(Integer.parseInt(props.get("codigofracao"))).getStringCellValue());
                        //quotaparte
                        fraction.setShare(currentRow.getCell(Integer.parseInt(props.get("quotaparte"))).getStringCellValue());

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
