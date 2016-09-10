package net.karrer.klionewsletter;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.nio.file.Path;
import java.nio.file.Paths;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;

import javafx.scene.layout.GridPane;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
 
@SuppressWarnings("restriction")
public final class KlioNewsletter extends Application {
 
    private static final String DEBUG_DIR =
    		(System.getProperty("os.name").equals("Linux")) ? "/home/karrer/klio-newsletter/data" : "c:/temp/klio";
    final TextArea msgArea = new TextArea();
    final Button saveButton = new Button("Save as ...");
    File docxFile;
    File csvfFile;
    File outfFile;
    Path tempFile;
 
    @Override
    public void start(final Stage stage) {
        stage.setTitle("KlioNewsletter");
        final Label title = new Label("Klio Newsletter");
        title.setFont(new Font("Arial", 24));
        final Button openDocxButton = new Button("Open Newsletter-Template (.docx)");
        final Button openCsvfButton = new Button("Open Artikel-Export-File  (.csv)");
        final TextField docxFileField = new TextField();
        final TextField csvfFileField = new TextField();
        final TextField outfFileField = new TextField();
        docxFileField.setMinWidth(500.0);
        
        // find the directory where the jar file is (except when debugging)
        String jarDir;
        try {
          jarDir = KlioNewsletter.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()
              .replaceFirst("^/(?=\\w:)", "").replaceFirst("/[^/]+$", "");
          if (jarDir.contains("target/classes")) {
            jarDir = DEBUG_DIR;
          }
        } catch (URISyntaxException e3) {
          jarDir = DEBUG_DIR;
        }
        File defaultDirectory = new File(jarDir);
        
 
        openDocxButton.setOnAction(
            new EventHandler<ActionEvent>() {
                @Override
                public void handle(final ActionEvent e) {
                  FileChooser fileChooser = new FileChooser();
                  fileChooser.setInitialDirectory(defaultDirectory);
                  fileChooser.getExtensionFilters().addAll(
                      new ExtensionFilter("Word docx Files", "*.docx"), new ExtensionFilter("All Files", "*.*"));
                    docxFile = fileChooser.showOpenDialog(stage);
                    if (docxFile != null) {
                      openFile("Template", docxFile, docxFileField, "read");
                      if (csvfFile != null) {
                        doConversion();
                      }
                    }
                }
            });
        
        openCsvfButton.setOnAction(
            new EventHandler<ActionEvent>() {
                @Override
                public void handle(final ActionEvent e) {
                  FileChooser fileChooser = new FileChooser();
                  fileChooser.setInitialDirectory(defaultDirectory);
                  fileChooser.getExtensionFilters().addAll(
                      new ExtensionFilter("Comma Separated Value files", "*.csv"), new ExtensionFilter("All Files", "*.*"));

                    csvfFile = fileChooser.showOpenDialog(stage);
                    if (csvfFile != null) {
                      openFile("Export", csvfFile, csvfFileField, "read");
                      if (docxFile != null) {
                        doConversion();
                      }
                    }
                }
            });

        saveButton.setOnAction(
            new EventHandler<ActionEvent>() {
                @Override
                public void handle(final ActionEvent e) {
                  FileChooser fileChooser = new FileChooser();
                  fileChooser.setInitialDirectory(defaultDirectory);
                  fileChooser.setTitle("Save File");
                  fileChooser.getExtensionFilters().addAll(
                      new ExtensionFilter("Word docx Files", "*.docx"));
                  File outfFile = fileChooser.showSaveDialog(stage);
                  if (outfFile != null) {
                      try {
                        Files.move(tempFile, Paths.get(outfFile.getPath()), StandardCopyOption.REPLACE_EXISTING);
                        openFile("Output", outfFile, outfFileField, "written");
                      } catch (IOException e1) {
                        msg("Error, cannot write " + outfFile.getPath()+"\n"+e+"\n");
                        System.err.println("Error, cannot write " + outfFile.getPath() + "\n" + e);
                      }
                    }
                }
            });
        
        msgArea.setMinSize(450.0, 450.0);
        saveButton.setDisable(true);

        final GridPane grid = new GridPane();
 
        GridPane.setConstraints(title, 0, 0, 2, 1);
        GridPane.setConstraints(openDocxButton, 0, 1);
        GridPane.setConstraints(docxFileField, 1, 1);
        GridPane.setConstraints(openCsvfButton, 0, 2);
        GridPane.setConstraints(csvfFileField, 1, 2);
        GridPane.setConstraints(msgArea, 0, 3, 2, 1);
        GridPane.setConstraints(saveButton, 0, 4);
        GridPane.setConstraints(outfFileField, 1, 4);
        
        grid.setHgap(6);
        grid.setVgap(6);
        grid.getChildren().addAll(title, openDocxButton, docxFileField, openCsvfButton, csvfFileField, msgArea, saveButton, outfFileField);
 
        final Pane rootGroup = new VBox(12);
        rootGroup.getChildren().addAll(grid);
        rootGroup.setPadding(new Insets(12, 12, 12, 12));
 
        stage.setScene(new Scene(rootGroup));
        stage.show();
    }
 
    public static void main(String[] args) {
        Application.launch(args);
    }
    
    public void msg(String mgs) {
      msgArea.appendText(mgs);
    }
 
    private void openFile(String desc, File file, TextField fld, String action) {
      msg(desc + " file "+ file.getPath() + " " + action + "\n");
      fld.setText(file.getPath());
    }
    
    private void doConversion() {
      EditDocx edoc = new EditDocx(docxFile, csvfFile);
      try {
        tempFile = edoc.convert();
      } catch (IOException e) {
        msg("Unspecific io exception" + e);
      }
      if (tempFile != null) {
        msg(edoc.getInfo());
        saveButton.setDisable(false);
      } else {
        msg(edoc.getError());
      }
    }

}