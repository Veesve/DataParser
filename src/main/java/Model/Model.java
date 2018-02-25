package Model;

import java.io.File;
import java.io.IOException;

public class Model {
    private File inputFile;
    private File outputFile;
    private Parser parser;

    public Model(File inputFile, File outputFile, Parser parser) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
        this.parser = parser;
    }

    public void parseData() throws IOException {
        parser.parseData(inputFile,outputFile);
    }
}
