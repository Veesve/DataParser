package Model;

import java.io.File;

public class ExcelParser implements Parser {
    @Override
    public void parseData(File inputFile, File outputFile) { //метод обработки данных из входного файла
        System.out.println(inputFile.getAbsolutePath());
        System.out.println(outputFile.getAbsolutePath());
    }
    private void createOutputFile(File outputFile){ //метод в котором должен создаваться XLSX файл в котором будет результат парса

    }
}
