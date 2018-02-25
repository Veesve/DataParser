package Model;

import java.io.File;
import java.io.IOException;

public interface Parser {
    //общий интерфейс для парсеров, на случай добавления новых файлов для применения паттерна "Стратегия"
    void parseData(File inputFile, File outputFile) throws IOException;
}
