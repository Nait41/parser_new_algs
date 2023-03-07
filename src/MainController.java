import data.InfoList;
import fileView.XLXSOpen;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.shape.Circle;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.ResourceBundle;

public class MainController {
    File loadFile;
    File saveFile;
    boolean checkLoad, checkUnload, checkStart = false;
    public static String errorMessageStr = "";

    @FXML
    private ResourceBundle resources;

    @FXML
    private URL location;

    @FXML
    private Button dirLoadButton;

    @FXML
    private Button dirUnloadButton;

    @FXML
    private Text loadStatus_end;

    @FXML
    private Text loadStatusFileNumber;

    @FXML
    private Button startButton;


    @FXML
    public Button closeButton;

    public MainController() throws IOException, InvalidFormatException {
    }

    public void addHinds(){

        Tooltip tipLoad = new Tooltip();
        tipLoad.setText("Выберите папку, в которой находится входной файл");
        tipLoad.setStyle("-fx-text-fill: turquoise;");
        dirLoadButton.setTooltip(tipLoad);

        Tooltip tipUnLoad = new Tooltip();
        tipUnLoad.setText("Выберите папку, в которой необходимо создать новый файл");
        tipUnLoad.setStyle("-fx-text-fill: turquoise;");
        dirUnloadButton.setTooltip(tipUnLoad);

        Tooltip tipStart = new Tooltip();
        tipStart.setText("Нажмите, для того, чтобы получить новый файл");
        tipStart.setStyle("-fx-text-fill: turquoise;");
        startButton.setTooltip(tipStart);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы выйти из программы");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

    }

    public void removeHinds(){
        dirLoadButton.setTooltip(null);
        dirUnloadButton.setTooltip(null);
        startButton.setTooltip(null);
        closeButton.setTooltip(null);
    }

    public static boolean tempHints = true;

    @FXML
    void initialize() throws IOException, InterruptedException, ClassNotFoundException {
        addHinds();

        FileInputStream loadStream = new FileInputStream(Application.rootDirPath + "\\load.png");
        Image loadImage = new Image(loadStream);
        ImageView loadView = new ImageView(loadImage);
        dirLoadButton.graphicProperty().setValue(loadView);

        FileInputStream unloadStream = new FileInputStream(Application.rootDirPath + "\\unload.png");
        Image unloadImage = new Image(unloadStream);
        ImageView unloadView = new ImageView(unloadImage);
        dirUnloadButton.graphicProperty().setValue(unloadView);

        FileInputStream startStream = new FileInputStream(Application.rootDirPath + "\\start.png");
        Image startImage = new Image(startStream);
        ImageView startView = new ImageView(startImage);
        startButton.graphicProperty().setValue(startView);

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);


        int r = 60;
        startButton.setShape(new Circle(r));
        startButton.setMinSize(r*2, r*2);
        startButton.setMaxSize(r*2, r*2);

        checkLoad = false;
        checkUnload = false;

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        dirLoadButton.setOnAction(actionEvent -> {
            if(!checkStart)
            {
                loadStatus_end.setText("");
                loadStatusFileNumber.setText("");
                FileChooser fileChooser = new FileChooser();
                File file = fileChooser.showOpenDialog(new Stage());
                loadFile = file;
                checkLoad = true;
            }
            else
            {
                errorMessageStr = "Происходит обработка файла. Повторите попытку попытку позже...";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

        dirUnloadButton.setOnAction(actionEvent -> {
                    if(!checkStart)
                    {
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        DirectoryChooser directoryChooser = new DirectoryChooser();
                        saveFile = directoryChooser.showDialog(new Stage());
                        checkUnload = true;
                    }
                    else
                    {
                        errorMessageStr = "Происходит обработка файла. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
        startButton.setOnAction(actionEvent -> {
                    if(!checkStart){
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        if(checkLoad & checkUnload){
                            checkStart = true;
                            new Thread(){
                                @Override
                                public void run(){
                                    MainLoader mainLoader = null;
                                    XLXSOpen xlxsOpen = null;
                                    InfoList infoList = null;
                                    OldAlgsOpen oldAlgsOpen = null;
                                    if(loadFile.getPath().contains(".xlsx"))
                                    {
                                        loadStatusFileNumber.setText("Обработка входного файла");
                                        try {
                                            oldAlgsOpen = new OldAlgsOpen(new File(Application.rootDirPath + "\\oldAlgs.xlsx"));
                                            xlxsOpen = new XLXSOpen(loadFile);
                                            mainLoader = new MainLoader(loadFile);
                                            infoList = new InfoList();
                                        } catch (IOException e) {
                                            e.printStackTrace();
                                        } catch (InvalidFormatException e) {
                                            e.printStackTrace();
                                        }
                                        try {
                                            xlxsOpen.getFileName(infoList);
                                            xlxsOpen.getBacteriaMediumRangeGenus(infoList);
                                            xlxsOpen.getBacteriaMediumRangeSpecies(infoList);
                                            oldAlgsOpen.getOldAlgs(infoList);
                                            mainLoader.setAllBacteria(infoList);
                                            mainLoader.setOldAlgs(infoList);
                                            mainLoader.setHighlightGenus();
                                            mainLoader.setHighlightSpecies();
                                            mainLoader.saveFile(saveFile);
                                            mainLoader.getClose();
                                            xlxsOpen.getClose();
                                        } catch (IOException e) {
                                            e.printStackTrace();
                                        }
                                    }
                                    loadStatusFileNumber.setText("");
                                    loadStatus_end.setText("Успешно сформирован новый файл!");
                                    checkStart = false;
                                }
                            }.start();
                        } else {
                            errorMessageStr = "Вы не указаали входной файл или директорию выгрузки...";
                            ErrorController errorController = new ErrorController();
                            try {
                                errorController.start(new Stage());
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    } else
                    {
                        errorMessageStr = "Происходит обработка файла. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
    }
}
