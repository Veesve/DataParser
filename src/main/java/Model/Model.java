package Model;

import java.io.File;
import java.io.IOException;

public class Model {
    private File inputFile;
    private File outputDir;
    private Parser parser;

    public Model(File inputFile, File outputDir, Parser parser) {
        this.inputFile = inputFile;
        this.outputDir = outputDir;
        this.parser = parser;
    }

    public void parseData() throws IOException {
        parser.parseData(inputFile, outputDir);
    }
}
