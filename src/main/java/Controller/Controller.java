package Controller;

import Model.Model;

public class Controller {
    private Model model;

    public void setModel(Model model) {
        this.model = model;
    }
    public void parseData(){
        model.parseData();
    }
}
