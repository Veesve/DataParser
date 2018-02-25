package Model;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;

public class ExcelParser implements Parser {

    @Override
    public void parseData(File inputFile, File outputFile) throws IOException { //метод обработки данных из входного файла

        System.out.println(inputFile.getAbsolutePath());
        System.out.println(outputFile.getAbsolutePath());

        NPOIFSFileSystem fs = new NPOIFSFileSystem(inputFile);                          // открываем файл, первый лист
        HSSFWorkbook workbook = new HSSFWorkbook(fs.getRoot(), true);
        HSSFSheet spreadsheet = workbook.getSheetAt(0);                          // TODO просмотр всех листов (групп)

        Row headRow = spreadsheet.getRow(6);                                  // строка для работы с оглавлением таблицы (строка 7)
        int headCellIndex = 1;                                                         // индекс для работы со столбцами оглавления

        int capacity = spreadsheet.getPhysicalNumberOfRows()-7;                        // получаем размер массива равный количеству студентов в группе

        ArrayList<String> marks = new ArrayList<String>(capacity);                          // строка оценок студента за все время
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

        // записываем в файл .csv с пометкой группы
        Writer wr = new OutputStreamWriter(new FileOutputStream(new File(outputFile
                                                                , inputFile.getName().replace(".xls",  "_" + spreadsheet.getSheetName() + ".csv")))
                                                                , StandardCharsets.UTF_8);
        /*Writer writer = new OutputStreamWriter(new FileOutputStream(outputFile +
                                                                    inputFile.getName().replace(".xls",  "_" + spreadsheet.getSheetName() + ".csv"))
                                                                    , StandardCharsets.UTF_8);*/
        for (String s : marks)
            wr.write(s+'\n');

        wr.close(); // сворачиваемся
        fs.close();

    }

    private void createOutputFile(File outputFile) { //метод в котором должен создаваться XLSX файл в котором будет результат парса

    }
}
