package Controller;

import Model.Model;

import java.io.IOException;

public class Controller {
    private Model model;

    public void setModel(Model model) {
        this.model = model;
    }
    public void parseData() throws IOException {
        model.parseData();
    }
}
