package org.openjfx.app;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.core.util.ZipUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import javafx.application.Application;
import javafx.concurrent.Task;
import javafx.concurrent.WorkerStateEvent;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.commons.collections4.CollectionUtils;
import org.openjfx.easyexcel.CustomStyleWriteHandler;
import org.openjfx.utils.file.FileUtil;
import org.openjfx.model.JobTask1ExcelModelRead2;
import org.openjfx.model.JobTask1ExcelModelRead3;
import org.openjfx.model.JobTask1ExcelModelWrite3;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * 合并发票
 */
public class InvoiceReportFileHandlerApp extends Application {

    private ProgressBar progressBar;
    private StackPane loadingPane;
    private Label fileLabel;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("处理Amazon广告发票文件");

        Button selectFileButton = new Button("选择文件");
        fileLabel = new Label("文件: ");

        selectFileButton.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                FileChooser fileChooser = new FileChooser();
                fileChooser.setTitle("打开文件");
                File selectedFile = fileChooser.showOpenDialog(primaryStage);
                if (selectedFile != null) {
                    fileLabel.setText("文件名: " + selectedFile.getName());
                    processFile(selectedFile);
                }
            }
        });

        progressBar = new ProgressBar();
        progressBar.setVisible(false);

        loadingPane = new StackPane();
        loadingPane.setStyle("-fx-background-color: rgba(0, 0, 0, 0.5)");
        loadingPane.setVisible(false);
        loadingPane.getChildren().add(new Label("Loading..."));

        VBox vBox = new VBox();
        vBox.setSpacing(10);
        vBox.setPadding(new Insets(10, 10, 10, 10));
        vBox.getChildren().addAll(selectFileButton, fileLabel, progressBar, loadingPane);

        primaryStage.setScene(new Scene(vBox, 400, 250));
        primaryStage.show();
    }

    private void processFile(File sourceFile) {
        progressBar.setVisible(true);
        loadingPane.setVisible(true);
        progressBar.setProgress(0);

        Task<Void> task = new Task<Void>() {
            @Override
            protected Void call() throws Exception {
                businessTask(sourceFile);
                return null;
            }
        };

        task.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                progressBar.setVisible(false);
                loadingPane.setVisible(false);
                displayResult("文件处理成功!");
                progressBar.progressProperty().unbind();
                progressBar.setProgress(1);
            }
        });
        task.setOnFailed(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                progressBar.setVisible(false);
                loadingPane.setVisible(false);
                displayResult("文件处理失败!");
                progressBar.progressProperty().unbind();
                progressBar.setProgress(1);
            }
        });

        progressBar.progressProperty().bind(task.progressProperty());

        new Thread(task).start();
    }

    private void displayResult(String result) {
        Stage dialog = new Stage();
        dialog.setTitle("处理结果");

        Label label = new Label(result);
        label.setStyle("-fx-font-size: 14pt; -fx-padding: 20px; -fx-text-alignment: center;");

        StackPane pane = new StackPane();
        pane.getChildren().add(label);

        Scene scene = new Scene(pane, 300, 100);
        dialog.setScene(scene);
        dialog.show();
    }

    public static void main(String[] args) {
        launch(args);
    }

    public void businessTask(File sourceFile) throws Exception {
        // 先解压文件
        String sourceRootFileName = sourceFile.getName();
        String sourceFilePrefixName = StrUtil.subBefore(sourceRootFileName, ".", true);
        String canonicalPath = sourceFile.getCanonicalPath();
        String sourceRootPath = StrUtil.subBefore(canonicalPath, File.separator, true);
        String targetFileConvertFolderPath = sourceRootPath + File.separator;
        String targetRootPath = targetFileConvertFolderPath + sourceFilePrefixName + "_target";
        File sourceRootFile = new File(targetRootPath);
        if (!sourceRootFile.exists()) {
            sourceRootFile.mkdirs();
        }
        ZipUtil.unzip(sourceFile, sourceRootFile);

        // 把所有符合格式的文件里面的数据，写到一个文件里面去
        ExcelWriter excelWriter = null;
        try {
            String tempFileName = sourceFilePrefixName + "_汇总所有数据" + System.currentTimeMillis() + ".xlsx";
            File tempFile = new File(targetFileConvertFolderPath + "/" + tempFileName);
            if (!tempFile.exists()) {
                tempFile.createNewFile();
            }
            excelWriter = EasyExcel.write(tempFile).build();
            WriteSheet writeSheet = EasyExcel.writerSheet(0, "汇总所有数据")
                    .registerWriteHandler(CustomStyleWriteHandler.buildDefaultWriteHandler(JobTask1ExcelModelWrite3.class))
                    .head(JobTask1ExcelModelWrite3.class).build();

            // 解析目标格式原始文件
            List<File> targetFileList = new ArrayList<>();
            FileUtil.listAllFile(sourceRootFile, targetFileList);
            for (File targetFile : targetFileList) {
                String sourceFileName = targetFile.getName();
                List<JobTask1ExcelModelRead2> targetFileInfoList = parseTargetFile(targetFile);
                if (CollectionUtils.isEmpty(targetFileInfoList)) {
                    continue;
                }
                System.out.println("原文件名：" + sourceFileName);
                // 重命名文件
                String targetInvoiceID = StrUtil.subBefore(sourceFileName, ".", false);
                List<JobTask1ExcelModelWrite3> writeList = new ArrayList<>();
                EasyExcel.read(targetFile, JobTask1ExcelModelRead3.class, new PageReadListener<JobTask1ExcelModelRead3>(dataList -> {
                    if (CollectionUtils.isNotEmpty(dataList)) {
                        for (JobTask1ExcelModelRead3 excelModelRead : dataList) {
                            if ("Total".equals(excelModelRead.getOrderDate()) || StrUtil.isEmpty(excelModelRead.getOrderDate())) {
                                continue;
                            }
                            JobTask1ExcelModelWrite3 excelModelWrite = new JobTask1ExcelModelWrite3();
                            BeanUtil.copyProperties(excelModelRead, excelModelWrite);
                            excelModelWrite.setInvoiceID(targetInvoiceID);
                            writeList.add(excelModelWrite);
                        }
                    }
                })).sheet(0).doRead();
                // 写入数据
                excelWriter.write(writeList, writeSheet);
                System.out.println("文件名：" + sourceFileName + "，解析并写入数据完成");
                System.out.println("原文件名：" + sourceFileName + "，解析结束");
            }
            System.out.println("所有符合条件的文件解析并写入数据完成");

        } catch (Exception e) {
            throw e;
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
            FileUtil.deleteWholeFile(targetRootPath);
        }

    }

    public static List<JobTask1ExcelModelRead2> parseTargetFile(File sourceFile) {
        List<JobTask1ExcelModelRead2> targetFileList = new ArrayList<>();
        System.out.println("解析文件名：" + sourceFile.getName());
        try {
            EasyExcel.read(sourceFile, JobTask1ExcelModelRead2.class, new PageReadListener<JobTask1ExcelModelRead2>(dataList -> {
                if (CollectionUtils.isNotEmpty(dataList)) {
                    targetFileList.addAll(dataList);
                }
            })).sheet(0).doRead();
        } catch (Exception e) {
            System.out.println("文件名：" + sourceFile.getName() + "，解析失败");
            e.printStackTrace();
        }

        return targetFileList;
    }
}
