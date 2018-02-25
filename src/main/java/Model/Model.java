package Model;

import java.io.File;

public class Model {
    private File inputFile;
    private File outputFile;
    private Parser parser;

    public Model(File inputFile, File outputFile, Parser parser) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
        this.parser = parser;
    }

    public void parseData(){
        parser.parseData(inputFile,outputFile);
    }
}
