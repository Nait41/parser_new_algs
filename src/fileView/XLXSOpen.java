package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class XLXSOpen {
    String fileName;
    Workbook workbook;
    public XLXSOpen(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        fileName = file.getName();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void getBacteriaMediumRangeGenus(InfoList infoList){
        for(int i = 1; i < workbook.getSheet("Genus").getPhysicalNumberOfRows(); i++){
            infoList.genus.add(new ArrayList<>());
            infoList.genus.get(i-1).add(workbook.getSheet("Genus").getRow(i).getCell(0).getStringCellValue());
            infoList.genus.get(i-1).add(Double.toString(workbook.getSheet("Genus").getRow(i).getCell(6).getNumericCellValue()));
            infoList.genus.get(i-1).add(Double.toString(workbook.getSheet("Genus").getRow(i).getCell(10).getNumericCellValue()));
        }
    }

    public void getBacteriaMediumRangeSpecies(InfoList infoList){
        for(int i = 1; i < workbook.getSheet("Species").getPhysicalNumberOfRows(); i++){
            infoList.species.add(new ArrayList<>());
            infoList.species.get(i-1).add(workbook.getSheet("Species").getRow(i).getCell(0).getStringCellValue());
            infoList.species.get(i-1).add(Double.toString(workbook.getSheet("Species").getRow(i).getCell(6).getNumericCellValue()));
            infoList.species.get(i-1).add(Double.toString(workbook.getSheet("Species").getRow(i).getCell(10).getNumericCellValue()));
        }
    }

    public void getFileName(InfoList infoList){
        infoList.fileName = fileName;
    }
}