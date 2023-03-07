import data.InfoList;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;

public class MainLoader extends JFrame {
    Workbook workbook;
    public MainLoader(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void setAllBacteria(InfoList infoList){
        ArrayList<ArrayList<String>> missedBacteria = new ArrayList<>();
        ArrayList<ArrayList<String>> fullBacteria = new ArrayList<>();
        CellStyle cellStyleFirst = workbook.getSheetAt(1).getRow(2).getCell(0).getCellStyle();
        CellStyle cellStyleSecond = workbook.getSheetAt(1).getRow(2).getCell(1).getCellStyle();
        for(int i = 0; i < infoList.genus.size(); i++){
            fullBacteria.add(new ArrayList<>());
            fullBacteria.get(fullBacteria.size() - 1).add(infoList.genus.get(i).get(0));
            fullBacteria.get(fullBacteria.size() - 1).add(infoList.genus.get(i).get(1));
            fullBacteria.get(fullBacteria.size() - 1).add(infoList.genus.get(i).get(2));
        }
        for(int i = 0; i < infoList.species.size(); i++){
            fullBacteria.add(new ArrayList<>());
            fullBacteria.get(fullBacteria.size() - 1).add(infoList.species.get(i).get(0));
            fullBacteria.get(fullBacteria.size() - 1).add(infoList.species.get(i).get(1));
            fullBacteria.get(fullBacteria.size() - 1).add(infoList.species.get(i).get(2));
        }
        boolean checkBacteria;
        boolean checkTemp;
        for (int k = 0; k < fullBacteria.size(); k++){
            checkBacteria = false;
            for(int i = 2; i < workbook.getSheetAt(1).getPhysicalNumberOfRows(); i++){
                if(workbook.getSheetAt(1).getRow(i).getCell(0).getStringCellValue().equals(fullBacteria.get(k).get(0))){
                    if (Double.parseDouble(fullBacteria.get(k).get(1)) == 0){
                        workbook.getSheetAt(1).getRow(i).createCell(3).setCellValue(0.0001);
                    } else {
                        workbook.getSheetAt(1).getRow(i).createCell(3).setCellValue(Double.parseDouble(fullBacteria.get(k).get(1)));
                    }

                    if (Double.parseDouble(fullBacteria.get(k).get(2)) == 0){
                        workbook.getSheetAt(1).getRow(i).createCell(4).setCellValue(0.01);
                    } else {
                        workbook.getSheetAt(1).getRow(i).createCell(4).setCellValue(Double.parseDouble(fullBacteria.get(k).get(2)));
                    }
                    workbook.getSheetAt(1).getRow(i).getCell(3).setCellStyle(cellStyleSecond);
                    workbook.getSheetAt(1).getRow(i).getCell(4).setCellStyle(cellStyleSecond);
                    checkBacteria = true;
                }
                if((i == workbook.getSheetAt(1).getPhysicalNumberOfRows() - 1) && !checkBacteria){
                    checkTemp = false;
                    for(int t = 0; t < missedBacteria.size();t++){
                        if(missedBacteria.get(t).get(0).equals(fullBacteria.get(k).get(0))){
                            checkTemp = true;
                        }
                    }
                    if(!checkTemp){
                        missedBacteria.add(new ArrayList<>());
                        missedBacteria.get(missedBacteria.size()-1).add(fullBacteria.get(k).get(0));
                        missedBacteria.get(missedBacteria.size()-1).add(fullBacteria.get(k).get(1));
                        missedBacteria.get(missedBacteria.size()-1).add(fullBacteria.get(k).get(2));
                    }
                }
            }
        }
        for(int i = workbook.getSheetAt(1).getPhysicalNumberOfRows(), k = 0; k < missedBacteria.size(); i++, k++){
            workbook.getSheetAt(1).createRow(i).createCell(0).setCellValue(missedBacteria.get(k).get(0));
            workbook.getSheetAt(1).getRow(i).createCell(1).setCellValue(0);

            if (Double.parseDouble(missedBacteria.get(k).get(1)) == 0){
                workbook.getSheetAt(1).getRow(i).createCell(3).setCellValue(0.0001);
            } else {
                workbook.getSheetAt(1).getRow(i).createCell(3).setCellValue(Double.parseDouble(fullBacteria.get(k).get(1)));
            }

            if (Double.parseDouble(missedBacteria.get(k).get(2)) == 0){
                workbook.getSheetAt(1).getRow(i).createCell(4).setCellValue(0.01);
            } else {
                workbook.getSheetAt(1).getRow(i).createCell(4).setCellValue(Double.parseDouble(fullBacteria.get(k).get(2)));
            }

            workbook.getSheetAt(1).getRow(i).getCell(0).setCellStyle(cellStyleFirst);
            workbook.getSheetAt(1).getRow(i).getCell(1).setCellStyle(cellStyleSecond);
            workbook.getSheetAt(1).getRow(i).getCell(3).setCellStyle(cellStyleSecond);
            workbook.getSheetAt(1).getRow(i).getCell(4).setCellStyle(cellStyleSecond);
        }
    }

    public void setHighlightGenus(){
        for(int i = 1; i < workbook.getSheet("Genus").getPhysicalNumberOfRows();i++){
            if(workbook.getSheet("Genus").getRow(i).getCell(13).getNumericCellValue() > 600){
                for (int j = 0; j < workbook.getSheet("Genus").getRow(i).getPhysicalNumberOfCells(); j++){
                    CellStyle cellStyle = workbook.getSheet("Genus").getWorkbook().createCellStyle();
                    cellStyle.setFillBackgroundColor(IndexedColors.BLACK.index);
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                    workbook.getSheet("Genus").getRow(i).getCell(j).setCellStyle(cellStyle);
                }
            }
        }
    }

    public void setHighlightSpecies(){
        for(int i = 1; i < workbook.getSheet("Species").getPhysicalNumberOfRows();i++){
            if(workbook.getSheet("Species").getRow(i).getCell(13).getNumericCellValue() > 500){
                for (int j = 0; j < workbook.getSheet("Species").getRow(i).getPhysicalNumberOfCells(); j++){
                    CellStyle cellStyle = workbook.getSheet("Species").getWorkbook().createCellStyle();
                    cellStyle.setFillBackgroundColor(IndexedColors.BLACK.index);
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                    workbook.getSheet("Species").getRow(i).getCell(j).setCellStyle(cellStyle);
                }
            }
        }
    }

    public void setOldAlgs(InfoList infoList){
        for(int i = 0; i < workbook.getSheetAt(1).getPhysicalNumberOfRows();i++){
            for (int j = 0; j < infoList.oldAlgs.size();j++){
                if(workbook.getSheetAt(1).getRow(i).getCell(0) != null && workbook.getSheetAt(1).getRow(i).getCell(0).getStringCellValue().equals(infoList.oldAlgs.get(j).get(0))){
                    if (infoList.oldAlgs.get(j).size() > 2
                            && workbook.getSheetAt(1).getRow(i).getCell(0) != null
                            && (workbook.getSheetAt(1).getRow(i).getCell(10) == null
                            || workbook.getSheetAt(1).getRow(i).getCell(10).getStringCellValue().equals(""))){
                        workbook.getSheetAt(1).getRow(i).createCell(10).setCellValue(infoList.oldAlgs.get(j).get(2));
                    }
                    if (infoList.oldAlgs.get(j).size() > 3){
                        if(infoList.oldAlgs.get(j).get(1).contains("отсутствует") || infoList.oldAlgs.get(j).get(1).contains("Отсутствует")){
                            workbook.getSheetAt(1).getRow(i).createCell(6).setCellValue(infoList.oldAlgs.get(j).get(3));
                        } else if(infoList.oldAlgs.get(j).get(1).contains("низкое") || infoList.oldAlgs.get(j).get(1).contains("Низкое")){
                            workbook.getSheetAt(1).getRow(i).createCell(7).setCellValue(infoList.oldAlgs.get(j).get(3));
                        } else if(infoList.oldAlgs.get(j).get(1).contains("среднее") || infoList.oldAlgs.get(j).get(1).contains("Среднее")){
                            workbook.getSheetAt(1).getRow(i).createCell(8).setCellValue(infoList.oldAlgs.get(j).get(3));
                        } else if(infoList.oldAlgs.get(j).get(1).contains("высокое") || infoList.oldAlgs.get(j).get(1).contains("Высокое")){
                            workbook.getSheetAt(1).getRow(i).createCell(9).setCellValue(infoList.oldAlgs.get(j).get(3));
                        }
                    }
                }
            }
        }
    }

    public void saveFile(File saveSample) throws IOException {
        workbook.write(new FileOutputStream(new File(saveSample.getPath() + "\\test.xlsx")));
        workbook.close();
    }
}

