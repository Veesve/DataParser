package View;

import Controller.Controller;
import Model.ExcelParser;
import Model.Model;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.beans.binding.BooleanBinding;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import java.io.File;
import java.io.IOException;

public class FXView extends Application {

    private static File inpXSLFile; //файл, получаемый для обработки данных
    private static File outXLSDir; //файл, в которой выводится результат обработки данных

    public void start(Stage primaryStage) {
        primaryStage.setTitle("Data Parser");
        primaryStage.resizableProperty().setValue(false);


        primaryStage.setOnCloseRequest(t -> {
            Platform.exit();
            System.exit(0);
        }); //закрытие процесса программы при завершении работы пользователя

        final FileChooser fileChooser = new FileChooser();

        final TextField inputFileName = new TextField();
        inputFileName.setPromptText("Укажите путь к excel файлу с данными студентов");
        inputFileName.setEditable(false);
        inputFileName.setPrefColumnCount(30);

        final Button inputFileButton = new Button("Выберите файл");

        inputFileButton.setOnAction(event -> {
            configureFileChooser(fileChooser);
            inpXSLFile= fileChooser.showOpenDialog(primaryStage);
            if(inpXSLFile!=null)
                inputFileName.setText(inpXSLFile.getAbsolutePath());
        }); //сохранение данных о файле в переменную и вывод пути в TextField

        final TextField outputFileName = new TextField();
        outputFileName.setPromptText("Укажите директорию  для файла с выходных данных");
        outputFileName.setText(fileChooser.getInitialFileName());
        outputFileName.setEditable(false);
        outputFileName.setPrefColumnCount(30);

        final Button outputFileButton = new Button("Укажите директорию");

        outputFileButton.setOnAction(event -> {
            final DirectoryChooser directoryChooser = new DirectoryChooser();
            outXLSDir = directoryChooser.showDialog(primaryStage);
            if (outXLSDir != null) {
                outputFileName.setText(outXLSDir.getAbsolutePath());
            }
        }); //сохранение данных о файле в переменную и вывод пути в TextField

        final Button confirmButton = new Button("Вывести данные в файл");
        BooleanBinding confirmButtonBinding = //проверка условия активации кнопки
                Bindings.and(inputFileName.textProperty().isEmpty(),outputFileName.textProperty().isEmpty());
        confirmButton.disableProperty().bind(confirmButtonBinding);
        confirmButton.setOnAction(event -> {
            Task parseData = new Task() { //специальный класс для потоков в JavaFX
                @Override
                protected Object call() throws IOException {
                    Controller controller = new Controller();
                    Model model = new Model(inpXSLFile, outXLSDir,new ExcelParser());
                    controller.setModel(model);
                    controller.parseData();
                    return null;
                }
            };
            new Thread(parseData).start();//создание нового потока, для того чтобы графический интерфейс
                                            // не "подвисал" при обработке данных
        });



        final GridPane inputGridPane = new GridPane();

        GridPane.setConstraints(inputFileName,0,0);
        GridPane.setConstraints(inputFileButton,1,0);
        GridPane.setConstraints(outputFileName,0,1);
        GridPane.setConstraints(outputFileButton,1,1);
        GridPane.setConstraints(confirmButton,1,2);

        inputGridPane.setHgap(6);
        inputGridPane.setVgap(6);
        inputGridPane.getChildren().addAll(inputFileName,inputFileButton,outputFileName,outputFileButton,confirmButton);

        inputFileButton.setMaxWidth(Double.MAX_VALUE); //установка ширины кнопок к единому значению
        outputFileButton.setMaxWidth(Double.MAX_VALUE); // не знаю почему этот код работает.
        confirmButton.setMaxWidth(Double.MAX_VALUE);// Взято из оракл доков https://docs.oracle.com/javafx/2/layout/size_align.htm


        final Pane rootGroup = new VBox(12);
        rootGroup.getChildren().addAll(inputGridPane);
        rootGroup.setPadding(new Insets(12,12,12,12));
        primaryStage.setScene(new Scene(rootGroup));
        primaryStage.show();

    }

    private static void configureFileChooser(final FileChooser fileChooser) { //метод, задающий простейшите настройки для FileChooser
        fileChooser.setTitle("Выберите excel файл");
        fileChooser.setInitialDirectory(
                new File(System.getProperty("user.home")) //установка начальной директории при выборе файла
        );
        fileChooser.getExtensionFilters().addAll( //установка фильтра на расширение
                new FileChooser.ExtensionFilter("XLS","*.xls"),
                new FileChooser.ExtensionFilter("XLSX","*.xlsx")
        );
    }

    public static void main(String[] args) {
        launch(args);
    }

}
