package Model;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;

public class ExcelParser implements Parser {

    @Override
    public void parseData(File inputFile, File outputDir) throws IOException { //метод обработки данных из входного файла

        NPOIFSFileSystem fs = new NPOIFSFileSystem(inputFile);                          // открываем файл, первый лист
        HSSFWorkbook workbook = new HSSFWorkbook(fs.getRoot(), true);
        HSSFSheet spreadsheet = workbook.getSheetAt(0);                          // TODO просмотр всех листов (групп)

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
       // createOutputCSVFile(inputFile, outputDir, fs, spreadsheet, marks);
        createOutputXLSFile(inputFile,outputDir,fs,spreadsheet,marks);


    }

    private void createOutputCSVFile(File inputFile,
                                     File outputDir,
                                     NPOIFSFileSystem fs,
                                     HSSFSheet spreadsheet,
                                     ArrayList<String> marks) throws IOException { //создание CSV файла со всеми выходными данными
        // записываем в файл .csv с пометкой группы
        String outputFileName = inputFile.getName().replace(".xls", "_" + spreadsheet.getSheetName() + ".csv");
        File outputFile = new File(outputDir, outputFileName);
        Writer wr = new OutputStreamWriter(new FileOutputStream(outputFile), StandardCharsets.UTF_8);

        for (String s : marks) {
            wr.write(s + ","+computeAverageMark(s)+'\n');
        }

        wr.close(); // сворачиваемся
        fs.close();
    }
    private void createOutputXLSFile(File inputFile,
                                     File outputDir,
                                     NPOIFSFileSystem fs,
                                     HSSFSheet spreadsheet,
                                     ArrayList<String> marks) throws IOException { //создание XLS файла с подсчитанным средним арифметическим
        String outputFileName = inputFile.getName().replace(".xls", "_" + spreadsheet.getSheetName() + ".xls");
        File outputFile = new File(outputDir,outputFileName);
        FileOutputStream fileOut = new FileOutputStream(outputFile);
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet worksheet = workbook.createSheet(spreadsheet.getSheetName());

        HSSFRow row1 = worksheet.createRow((short)0);

        HSSFCell cellA1 = row1.createCell((short)0);
        cellA1.setCellValue("ФИО студента");
        HSSFCell cellB1 = row1.createCell((short)1);
        cellB1.setCellValue("Средний балл");
        for(int i = 0;i<marks.size();i++){
            HSSFRow rowI = worksheet.createRow((short)i+1);
            String s = marks.get(i);
            HSSFCell cell1 = rowI.createCell((short)0);
            cell1.setCellValue(s.split(",")[0]);

            HSSFCell cell2 = rowI.createCell((short)1);
            cell2.setCellValue(computeAverageMark(s));
        }
        worksheet.autoSizeColumn(0);
        worksheet.autoSizeColumn(1);
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
        return (double)sum/count;
    }

}
