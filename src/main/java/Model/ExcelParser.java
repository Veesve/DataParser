package Model;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelParser implements Parser {

    @Override
    public void parseData(File inputFile, File outputDir) throws IOException { //метод обработки данных из входного файла

        NPOIFSFileSystem fs = new NPOIFSFileSystem(inputFile);                          // открываем файл, первый лист
        HSSFWorkbook workbook = new HSSFWorkbook(fs.getRoot(), true);
        List<String> spreadsheetNames = new ArrayList<>(); //список имен таблиц
        List<List<String>> marksList = new ArrayList<>();//список оценок студентов

        int sheetsCount = workbook.getNumberOfSheets();
        for(int i = 0;i<sheetsCount;i++) {
            HSSFSheet spreadsheet = workbook.getSheetAt(i);
            spreadsheetNames.add(spreadsheet.getSheetName());
            ArrayList<String> marks = getSheetInfo(spreadsheet);
            marksList.add(marks);
        }
        createOutputXLSFile(outputDir,fs,spreadsheetNames,marksList);


    }

    private ArrayList<String> getSheetInfo(HSSFSheet spreadsheet) { //получение строки данных из одного листа
        Row headRow = spreadsheet.getRow(6);                                  // строка для работы с оглавлением таблицы (строка 7)
        int headCellIndex = 1;                                                         // индекс для работы со столбцами оглавления

        int capacity = spreadsheet.getPhysicalNumberOfRows()-7;                        // получаем размер массива равный количеству студентов в группе

        ArrayList<String> marks = new ArrayList<>(capacity);                          // строка оценок студента за все время
        for (int i = 1; i <= capacity; i++)
            marks.add(spreadsheet.getRow(6+i).getCell(1).getStringCellValue()); // заполняем ФИО

        while (headCellIndex < headRow.getLastCellNum()-1) {                           // пока есть столбцы
            String headValue = headRow.getCell(++headCellIndex).getStringCellValue();  // проверяем на соответствие regex для Экзаменов и т.д.
            if (headValue.matches(".*(Э|КР|КП|Практика.*)")) {
                Row bodyRow = spreadsheet.getRow(7);                          // первая строка для работы со студентами
                int bodyRowIndex = 8;                                                  // индекс для работы со строками студентов
                for (int i = 0; i < capacity; i++) {                                   // заполняем данный Экзамен для всех студентов
                    String temp = bodyRow.getCell(headCellIndex).getStringCellValue();
                    if (temp.equals("У"))
                        temp = ",3";
                    else if (temp.equals("Х"))
                        temp = ",4";
                    else if (temp.equals("О"))
                        temp = ",5";
                    else if (bodyRow.getCell(headCellIndex).getStringCellValue().equals("")) // выход, если находим НЕНАСТУПИВШИЙ экзамен
                        break;
                    else temp = ",";
                    marks.set(i, marks.get(i)+temp);
                    bodyRow = spreadsheet.getRow(bodyRowIndex++);
                }
            }
        }
        return marks;
    }

    private void createOutputXLSFile( File outputDir,
                                     NPOIFSFileSystem fs,
                                     List<String> spreadsheetNames,
                                     List<List<String>> marksList) throws IOException { //создание XLS файла с подсчитанным средним арифметическим
        File outputFile = new File(outputDir,"Date Analyse Result.xls");
        FileOutputStream fileOut = new FileOutputStream(outputFile);
        HSSFWorkbook workbook = new HSSFWorkbook();
        for(int i = 0;i<spreadsheetNames.size();i++) { //воссоздание страниц для каждой группы отдельно
            HSSFSheet worksheet = workbook.createSheet(spreadsheetNames.get(i));
            List<String> marks = marksList.get(i);
            HSSFRow row1 = worksheet.createRow((short) 0);

            HSSFCell cellA1 = row1.createCell((short) 0);
            cellA1.setCellValue("ФИО студента");
            HSSFCell cellB1 = row1.createCell((short) 1);
            cellB1.setCellValue("Средний балл");
            for (int j = 0; j < marks.size(); j++) {
                HSSFRow rowI = worksheet.createRow((short) j + 1);
                String s = marks.get(j);
                HSSFCell cell1 = rowI.createCell((short) 0);
                cell1.setCellValue(s.split(",")[0]);

                HSSFCell cell2 = rowI.createCell((short) 1);
                cell2.setCellValue(computeAverageMark(s));
            }
            worksheet.autoSizeColumn(0);
            worksheet.autoSizeColumn(1);
        }
        workbook.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }


    private double computeAverageMark(String csvString){ //метод подсчёта средней арифметической оценки из CSV строки
        int sum = 0;
        int count = 0;
        String[] csvValues = csvString.split(",");
        for(int i = 1;i<csvValues.length;i++){
            switch (csvValues[i]){
                case "3":{
                    sum+=3;
                    count++;
                    break;
                }
                case "4":{
                    sum+=3;
                    count++;
                }
                case "5":{
                    sum+=4;
                    count++;
                }
                default:{ }
            }
        }
        return round((double)sum/count,2);
    }

    private static double round(double value, int places) { //округление числа до необходимого количества знаков после запятой
        if(places < 0) throw new IllegalArgumentException();

        long factor = (long)Math.pow(10,places);
        value = value*factor;
        long tmp = Math.round(value);
        return (double)tmp/factor;
    }

}
