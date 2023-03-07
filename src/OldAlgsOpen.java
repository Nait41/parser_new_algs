import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class OldAlgsOpen {

    Workbook workbook;
    public OldAlgsOpen(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void getOldAlgs(InfoList infoList){
        for (int i = 0 ; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++){
            if(workbook.getSheetAt(0).getRow(i).getCell(0) != null && !workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue().equals("")){
                infoList.oldAlgs.add(new ArrayList<>());
                infoList.oldAlgs.get(infoList.oldAlgs.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
            }
            if(workbook.getSheetAt(0).getRow(i).getCell(2) != null && !workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue().equals("")){
                infoList.oldAlgs.get(infoList.oldAlgs.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue());
            }
            if(workbook.getSheetAt(0).getRow(i).getCell(3) != null && !workbook.getSheetAt(0).getRow(i).getCell(3).getStringCellValue().equals("")){
                infoList.oldAlgs.get(infoList.oldAlgs.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
            }
            if(workbook.getSheetAt(0).getRow(i).getCell(4) != null && !workbook.getSheetAt(0).getRow(i).getCell(4).getStringCellValue().equals("")){
                infoList.oldAlgs.get(infoList.oldAlgs.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(4).getStringCellValue());
            }
        }
    }

    public void getClose() throws IOException {
        workbook.close();
    }


}
