package org.openjfx.app;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import javafx.application.Application;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.paint.Paint;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.openjfx.easyexcel.CustomStyleWriteHandler;
import org.openjfx.model.JobTask1ExcelModelRead;
import org.openjfx.model.JobTask1ExcelModelRead2;
import org.openjfx.model.JobTask1ExcelModelRead3;
import org.openjfx.model.JobTask1ExcelModelWrite3;
import org.openjfx.utils.file.FileUtil;

import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 匹配发票并合并发票v3
 * 新增支持：拖拽文件上传，鼠标指针悬停显示完整文件名；
 * 优化UI界面
 */
@Slf4j
public class InvoiceReportFileHandlerAppV3 extends Application {

    private File contraCogsInvoicesSourceFile;
    private File zipInvoicesSourceFile;
    private Label contraCogsInvoicesSourceFileLabel;
    private Label zipInvoicesSourceFileLabel;
    private Label contraCogsInvoicesSourceFileErrorLabel;
    private Label zipInvoicesSourceFileErrorLabel;
    private Button submitButton;
    private Button resetButton;

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        // 初始化控件
        Label label1 = new Label("ContraCogsInvoices:");
        Label label2 = new Label("待处理发票压缩包:");

        Button uploadButton1 = new Button("上传");
        Button uploadButton2 = new Button("上传");

        contraCogsInvoicesSourceFileLabel = new Label("拖拽文件到此处");
        zipInvoicesSourceFileLabel = new Label("拖拽文件到此处");

        contraCogsInvoicesSourceFileErrorLabel = new Label();
        contraCogsInvoicesSourceFileErrorLabel.setTextFill(Color.RED);
        zipInvoicesSourceFileErrorLabel = new Label();
        zipInvoicesSourceFileErrorLabel.setTextFill(Color.RED);

        submitButton = new Button("提交");
        resetButton = new Button("重置");

        // 美化标签
        label1.setStyle("-fx-font-size: 16px; -fx-font-weight: bold;");
        label2.setStyle("-fx-font-size: 16px; -fx-font-weight: bold;");

        // 修改按钮颜色
        uploadButton1.setStyle("-fx-background-color: #3498DB; -fx-text-fill: white; -fx-font-size: 14px;");
        uploadButton2.setStyle("-fx-background-color: #3498DB; -fx-text-fill: white; -fx-font-size: 14px;");
        submitButton.setStyle("-fx-background-color: #2ECC71; -fx-text-fill: white; -fx-font-size: 16px; -fx-pref-width: 120px;");
        resetButton.setStyle("-fx-background-color: #E74C3C; -fx-text-fill: white; -fx-font-size: 16px; -fx-pref-width: 120px;");

        // 提示标签样式
        contraCogsInvoicesSourceFileLabel.setFont(Font.font("Arial", FontWeight.BOLD, 14));
        contraCogsInvoicesSourceFileLabel.setTextFill(Color.GRAY);
        contraCogsInvoicesSourceFileLabel.setPadding(new Insets(10));
        contraCogsInvoicesSourceFileLabel.setAlignment(Pos.CENTER);
        contraCogsInvoicesSourceFileLabel.setBorder(new Border(new BorderStroke(Color.GRAY, BorderStrokeStyle.DASHED,
                new CornerRadii(5), new BorderWidths(2))));

        zipInvoicesSourceFileLabel.setFont(Font.font("Arial", FontWeight.BOLD, 14));
        zipInvoicesSourceFileLabel.setTextFill(Color.GRAY);
        zipInvoicesSourceFileLabel.setPadding(new Insets(10));
        zipInvoicesSourceFileLabel.setAlignment(Pos.CENTER);
        zipInvoicesSourceFileLabel.setBorder(new Border(new BorderStroke(Color.GRAY, BorderStrokeStyle.DASHED,
                new CornerRadii(5), new BorderWidths(2))));

        // 设置上传按钮1的点击事件
        uploadButton1.setOnAction(e -> {
            contraCogsInvoicesSourceFile = handleFileSelection(primaryStage);
            if (contraCogsInvoicesSourceFile != null) {
                updateDragAndDropFileLabel(contraCogsInvoicesSourceFileLabel, contraCogsInvoicesSourceFile.getName(), Color.GREEN);
                contraCogsInvoicesSourceFileErrorLabel.setText("");
            }
        });

        // 设置上传按钮2的点击事件
        uploadButton2.setOnAction(e -> {
            zipInvoicesSourceFile = handleFileSelection(primaryStage);
            if (zipInvoicesSourceFile != null) {
                updateDragAndDropFileLabel(zipInvoicesSourceFileLabel, zipInvoicesSourceFile.getName(), Color.GREEN);
                zipInvoicesSourceFileErrorLabel.setText("");
            }
        });

        // 设置提交按钮的点击事件
        submitButton.setOnAction(e -> handleSubmit());

        // 设置重置按钮的点击事件
        resetButton.setOnAction(e -> handleReset());

        // 配置拖拽文件上传功能
        configureDragAndDrop(contraCogsInvoicesSourceFileLabel, 1);
        configureDragAndDrop(zipInvoicesSourceFileLabel, 2);

        // 布局设置
        GridPane gridPane = new GridPane();
        gridPane.setHgap(20);
        gridPane.setVgap(10);
        gridPane.setPadding(new Insets(20, 30, 20, 30));

        gridPane.add(label1, 0, 0);
        gridPane.add(uploadButton1, 1, 0);
        gridPane.add(contraCogsInvoicesSourceFileLabel, 2, 0);
        gridPane.add(contraCogsInvoicesSourceFileErrorLabel, 1, 1, 2, 1); // 添加错误提示标签

        gridPane.add(label2, 0, 2);
        gridPane.add(uploadButton2, 1, 2);
        gridPane.add(zipInvoicesSourceFileLabel, 2, 2);
        gridPane.add(zipInvoicesSourceFileErrorLabel, 1, 3, 2, 1); // 添加错误提示标签

        // 提交和重置按钮布局
        HBox buttonBox = new HBox(20, submitButton, resetButton);
        buttonBox.setAlignment(Pos.CENTER);

        VBox mainLayout = new VBox(30, gridPane, buttonBox);
        mainLayout.setAlignment(Pos.CENTER);
        mainLayout.setStyle("-fx-background-color: #ECF0F1; -fx-padding: 40px;");

        Scene scene = new Scene(mainLayout, 700, 400);
        primaryStage.setTitle("发票小工具v3");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private File handleFileSelection(Stage primaryStage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("选择文件");
        return fileChooser.showOpenDialog(primaryStage);
    }

    private void handleSubmit() {
        boolean valid = true;

        if (contraCogsInvoicesSourceFile == null) {
            contraCogsInvoicesSourceFileErrorLabel.setText("* 此项必填");
            valid = false;
        } else {
            contraCogsInvoicesSourceFileErrorLabel.setText("");
        }

        if (zipInvoicesSourceFile == null) {
            zipInvoicesSourceFileErrorLabel.setText("* 此项必填");
            valid = false;
        } else {
            zipInvoicesSourceFileErrorLabel.setText("");
        }

        if (zipInvoicesSourceFile != null) {
            String zipInvoicesSourceFileName = zipInvoicesSourceFile.getName();
            String zipInvoicesSourceFileSuffixName = StrUtil.subAfter(zipInvoicesSourceFileName, ".", true);
            if (!"zip".equalsIgnoreCase(zipInvoicesSourceFileSuffixName)) {
                zipInvoicesSourceFileErrorLabel.setText("仅限上传zip压缩包");
                zipInvoicesSourceFileErrorLabel.setTextFill(Color.RED);
                valid = false;
            }
        }


        if (valid) {
            submitButton.setDisable(true);
            submitButton.setText("处理中...");

            Task<Void> task = new Task<Void>() {
                @Override
                protected Void call() throws Exception {
                    processFiles(contraCogsInvoicesSourceFile, zipInvoicesSourceFile);
                    return null;
                }

                @Override
                protected void succeeded() {
                    submitButton.setDisable(false);
                    submitButton.setText("提交");
                    showAlert(Alert.AlertType.INFORMATION, "成功", "处理成功");
                }

                @Override
                protected void failed() {
                    submitButton.setDisable(false);
                    submitButton.setText("提交");
                    showAlert(Alert.AlertType.ERROR, "失败", "处理失败");
                }
            };

            new Thread(task).start();
        }
    }

    private void handleReset() {
        contraCogsInvoicesSourceFile = null;
        zipInvoicesSourceFile = null;
        contraCogsInvoicesSourceFileErrorLabel.setText("");
        zipInvoicesSourceFileErrorLabel.setText("");

        // 提示标签样式
        contraCogsInvoicesSourceFileLabel.setText("拖拽文件到此处");
        contraCogsInvoicesSourceFileLabel.setFont(Font.font("Arial", FontWeight.BOLD, 14));
        contraCogsInvoicesSourceFileLabel.setTextFill(Color.GRAY);
        contraCogsInvoicesSourceFileLabel.setPadding(new Insets(10));
        contraCogsInvoicesSourceFileLabel.setAlignment(Pos.CENTER);
        contraCogsInvoicesSourceFileLabel.setBorder(new Border(new BorderStroke(Color.GRAY, BorderStrokeStyle.DASHED,
                new CornerRadii(5), new BorderWidths(2))));
        zipInvoicesSourceFileLabel.setText("拖拽文件到此处");
        zipInvoicesSourceFileLabel.setFont(Font.font("Arial", FontWeight.BOLD, 14));
        zipInvoicesSourceFileLabel.setTextFill(Color.GRAY);
        zipInvoicesSourceFileLabel.setPadding(new Insets(10));
        zipInvoicesSourceFileLabel.setAlignment(Pos.CENTER);
        zipInvoicesSourceFileLabel.setBorder(new Border(new BorderStroke(Color.GRAY, BorderStrokeStyle.DASHED,
                new CornerRadii(5), new BorderWidths(2))));
    }

    private void configureDragAndDrop(Label label, int labelNumber) {
        label.setOnDragOver(event -> {
            if (event.getGestureSource() != label && event.getDragboard().hasFiles()) {
                event.acceptTransferModes(TransferMode.COPY);
            }
            event.consume();
        });

        label.setOnDragDropped(event -> {
            Dragboard db = event.getDragboard();
            boolean success = false;
            if (db.hasFiles()) {
                success = true;
                File file = db.getFiles().get(0);
                if (labelNumber == 1) {
                    contraCogsInvoicesSourceFile = file;
                    updateDragAndDropFileLabel(contraCogsInvoicesSourceFileLabel, contraCogsInvoicesSourceFile.getName(), Color.GREEN);
                    contraCogsInvoicesSourceFileErrorLabel.setText(""); // 清除错误信息
                } else if (labelNumber == 2) {
                    zipInvoicesSourceFile = file;
                    updateDragAndDropFileLabel(zipInvoicesSourceFileLabel, zipInvoicesSourceFile.getName(), Color.GREEN);
                    zipInvoicesSourceFileErrorLabel.setText(""); // 清除错误信息
                }
            }
            event.setDropCompleted(success);
            event.consume();
        });
    }

    private void updateDragAndDropFileLabel(Label label, String text, Paint paint) {
        int maxLength = 30; // 设置显示的最大字符长度
        String displayName = text;
        if (text.length() > maxLength) {
            displayName = text.substring(0, maxLength - 3) + "..."; // 超出部分用省略号代替
        }
        label.setText(displayName);
        label.setTextFill(paint);
        label.setBorder(new Border(new BorderStroke(paint, BorderStrokeStyle.SOLID,
                new CornerRadii(5), new BorderWidths(2))));

        // 设置Tooltip用于显示完整文件名
        Tooltip tooltip = new Tooltip(text);
        Tooltip.install(label, tooltip);
    }

    private void showAlert(Alert.AlertType alertType, String title, String message) {
        Alert alert = new Alert(alertType);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

    /**
     * 处理文件
     *
     * @param contraCogsInvoicesSourceFile
     * @param zipInvoicesSourceFile
     * @return
     */
    public boolean processFiles(File contraCogsInvoicesSourceFile, File zipInvoicesSourceFile) {
        try {
            System.out.println("contraCogsInvoicesSourceFile:" + contraCogsInvoicesSourceFile.getName());
            System.out.println("zipInvoicesSourceFile:" + zipInvoicesSourceFile.getName());

            // 先解压文件
            String zipInvoicesSourceFileName = zipInvoicesSourceFile.getName();
            String zipInvoicesNewFolderName = StrUtil.subBefore(zipInvoicesSourceFileName, ".", true) + System.currentTimeMillis();
            String zipInvoicesSourceFileCanonicalPath = zipInvoicesSourceFile.getCanonicalPath();
            String zipInvoicesSourceFileRootPath = StrUtil.subBefore(zipInvoicesSourceFileCanonicalPath, File.separator, true);
            String zipInvoicesSourceFileUnZipPath = zipInvoicesSourceFileRootPath + File.separator + zipInvoicesNewFolderName + "_Unzip";
            File zipInvoicesSourceFileUnZipPathFile = new File(zipInvoicesSourceFileUnZipPath);
            if (!zipInvoicesSourceFileUnZipPathFile.exists()) {
                zipInvoicesSourceFileUnZipPathFile.mkdirs();
            }
            FileUtil.uncompressAllFile(zipInvoicesSourceFile, zipInvoicesSourceFileUnZipPath);
            // 创建金额重命名文件夹
            String zipInvoicesSourceFileUnZipAddAmountPath = zipInvoicesSourceFileRootPath + File.separator + zipInvoicesNewFolderName + "_AddAmount";
            File zipInvoicesSourceFileUnZipAddAmountPathFile = new File(zipInvoicesSourceFileUnZipAddAmountPath);
            if (!zipInvoicesSourceFileUnZipAddAmountPathFile.exists()) {
                zipInvoicesSourceFileUnZipAddAmountPathFile.mkdirs();
            }
            // 创建Invoice ID重命名文件夹
            String zipInvoicesSourceFileUnZipRenameInvoiceIDPath = zipInvoicesSourceFileRootPath + File.separator + zipInvoicesNewFolderName + "_InvoiceID";
            File zipInvoicesSourceFileUnZipRenameInvoiceIDPathFile = new File(zipInvoicesSourceFileUnZipRenameInvoiceIDPath);
            if (!zipInvoicesSourceFileUnZipRenameInvoiceIDPathFile.exists()) {
                zipInvoicesSourceFileUnZipRenameInvoiceIDPathFile.mkdirs();
            }

            // 匹配可用文件，并用金额重命名文件
            this.matchAvailableFileAndRenameAmount(zipInvoicesSourceFileUnZipPath, zipInvoicesSourceFileUnZipAddAmountPath);

            // 金额重命名文件，匹配contraCogsInvoicesSourceFile，用Invoice ID重命名文件，并移动到新文件夹
            this.renameAndRemoveAvailableFile(contraCogsInvoicesSourceFile, zipInvoicesSourceFileUnZipAddAmountPath, zipInvoicesSourceFileUnZipRenameInvoiceIDPath);

            // 合并Invoice ID可用文件
            List<File> availableInvoiceIDFileList = new ArrayList<>();
            FileUtil.listAllFile(zipInvoicesSourceFileUnZipRenameInvoiceIDPathFile, availableInvoiceIDFileList);
            if (CollectionUtils.isNotEmpty(availableInvoiceIDFileList)) {
                String mergeAvailablePathFileName = zipInvoicesSourceFileRootPath + File.separator + zipInvoicesNewFolderName + "_汇总InvoiceID数据.xlsx";
                File mergeAvailablePathFile = new File(mergeAvailablePathFileName);
                if (!mergeAvailablePathFile.exists()) {
                    mergeAvailablePathFile.createNewFile();
                }
                this.mergeAvailableFile(mergeAvailablePathFile, availableInvoiceIDFileList);
            }

            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    /**
     * 匹配可用文件，并用金额重命名文件
     *
     * @param zipInvoicesSourceFileUnZipPath
     * @param zipInvoicesSourceFileUnZipAddAmountPath
     * @throws Exception
     */
    public void matchAvailableFileAndRenameAmount(String zipInvoicesSourceFileUnZipPath, String zipInvoicesSourceFileUnZipAddAmountPath) throws Exception {
        File zipInvoicesSourceFileUnZipFile = new File(zipInvoicesSourceFileUnZipPath);
        List<File> unZipFileList = new ArrayList<>();
        FileUtil.listAllFile(zipInvoicesSourceFileUnZipFile, unZipFileList);
        for (File sourceFile : unZipFileList) {
            List<JobTask1ExcelModelRead2> sourceFileInfoList = parseTargetFile(sourceFile);
            if (CollectionUtils.isEmpty(sourceFileInfoList)) {
                continue;
            }
            Long notEmptyCount = sourceFileInfoList.stream().filter(v -> StrUtil.isNotEmpty(v.getRebateInAgreementCurrency()) && StrUtil.isNotEmpty(v.getOrderDate())).collect(Collectors.counting());
            if (notEmptyCount == 0) {
                continue;
            }

            // 获取Total数据
            BigDecimal numRebateInAgreementCurrency = null;
            try {
                numRebateInAgreementCurrency = sourceFileInfoList.stream()
                        .filter(v -> !"Total".equals(v.getOrderDate()))
                        .map(v -> v.getRebateInAgreementCurrency())
                        .collect(Collectors.reducing(BigDecimal.ZERO, v -> getRebateInAgreementCurrencyNum(v), BigDecimal::add))
                        .setScale(2, BigDecimal.ROUND_HALF_UP);
            } catch (Exception e) {
                log.error("原文件名：{}，解析计算Total失败", sourceFile.getName());
                continue;
            }
            log.info("Total：{}， 原文件名：{}", numRebateInAgreementCurrency.toString(), sourceFile.getName());

            // 重命名文件
            String suffixName = StrUtil.subAfter(sourceFile.getName(), ".", true);
            String preName = StrUtil.subBefore(sourceFile.getName(), ".", true);
            String renameFileName = numRebateInAgreementCurrency.toString() + "_" + preName + "." + suffixName;
            boolean renameToFlag = sourceFile.renameTo(new File(zipInvoicesSourceFileUnZipAddAmountPath + File.separator + renameFileName));
            log.info("重命名文件：{}，{}", (zipInvoicesSourceFileUnZipAddAmountPath + File.separator + renameFileName), renameToFlag ? "成功" : "失败");
        }
    }

    public static List<JobTask1ExcelModelRead2> parseTargetFile(File sourceFile) {
        List<JobTask1ExcelModelRead2> targetFileList = new ArrayList<>();
        String sourceFileName = sourceFile.getName();
        String sourceFileNameSuffix = StrUtil.subAfter(sourceFileName, ".", true);
        if (!"xls".equals(sourceFileNameSuffix) && !"xlsx".equals(sourceFileNameSuffix)) {
            return targetFileList;
        }
        log.info("解析文件名：{}", sourceFile.getName());
        try {
            EasyExcel.read(sourceFile, JobTask1ExcelModelRead2.class, new PageReadListener<JobTask1ExcelModelRead2>(dataList -> {
                if (CollectionUtils.isNotEmpty(dataList)) {
                    targetFileList.addAll(dataList);
                }
            })).sheet(0).doRead();
        } catch (Exception e) {
            log.info("解析文件名：{}, 解析失败", sourceFile.getName());
            e.printStackTrace();
        }

        return targetFileList;
    }

    public static BigDecimal getRebateInAgreementCurrencyNum(String value) {
        if (StrUtil.isEmpty(value)) {
            return BigDecimal.ZERO;
        }

        try {
            value = value.replace(",", "");
            if (value.startsWith("-")) {
                value = value.replace("-", "");
                return new BigDecimal(value).negate();
            }

            return new BigDecimal(value);
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
    }

    /**
     * 金额重命名文件，匹配contraCogsInvoicesSourceFile，用Invoice ID重命名文件，并移动到新文件夹
     *
     * @param contraCogsInvoicesSourceFile
     * @param zipInvoicesSourceFileUnZipAddAmountPath
     * @param zipInvoicesSourceFileUnZipRenameInvoiceIDPath
     * @throws Exception
     */
    public void renameAndRemoveAvailableFile(File contraCogsInvoicesSourceFile, String zipInvoicesSourceFileUnZipAddAmountPath,
                                             String zipInvoicesSourceFileUnZipRenameInvoiceIDPath) throws Exception {
        List<JobTask1ExcelModelRead> contraCogsInvoicesList = parseContraCogsInvoices(contraCogsInvoicesSourceFile);
        System.out.println(contraCogsInvoicesList.size());

        HashMap<String, List<JobTask1ExcelModelRead>> contraCogsInvoicesGroupMap = contraCogsInvoicesList.stream()
                .collect(Collectors.groupingBy(v -> v.getOriginalBalance(), HashMap::new, Collectors.toList()));
        List<JobTask1ExcelModelRead> contraCogsInvoicesDistinctList = contraCogsInvoicesGroupMap.entrySet().stream()
                .filter(v -> v.getValue().size() == 1)
                .map(v -> v.getValue().get(0)).collect(Collectors.toList());

        Map<BigDecimal, String> contraCogsInvoicesInfoMap = contraCogsInvoicesDistinctList.stream()
                .collect(Collectors.toMap(a -> getOriginalBalanceNum(a.getOriginalBalance()), b -> b.getInvoiceID(), (v1, v2) -> v2));
        System.out.println(contraCogsInvoicesInfoMap.toString());

        // 列出所有金额重命名文件
        List<File> unZipAddAmountFileList = new ArrayList<>();
        FileUtil.listAllFile(new File(zipInvoicesSourceFileUnZipAddAmountPath), unZipAddAmountFileList);
        for (File convertAddAmountFile : unZipAddAmountFileList) {
            List<JobTask1ExcelModelRead2> convertAddAmountFileInfoList = parseTargetFile(convertAddAmountFile);
            if (CollectionUtils.isEmpty(convertAddAmountFileInfoList)) {
                continue;
            }

            // 获取Total数据
            BigDecimal numRebateInAgreementCurrency = null;
            try {
                numRebateInAgreementCurrency = convertAddAmountFileInfoList.stream()
                        .filter(v -> !"Total".equals(v.getOrderDate()))
                        .map(v -> v.getRebateInAgreementCurrency())
                        .collect(Collectors.reducing(BigDecimal.ZERO, v -> getRebateInAgreementCurrencyNum(v), BigDecimal::add))
                        .setScale(2, BigDecimal.ROUND_HALF_UP);
            } catch (Exception e) {
                log.error("原文件名：{}，解析计算Total失败", convertAddAmountFile.getName());
                continue;
            }

            log.info("Total：{}， 原文件名：{}", numRebateInAgreementCurrency.toString(), convertAddAmountFile.getName());

            for (Map.Entry<BigDecimal, String> entry : contraCogsInvoicesInfoMap.entrySet()) {
                // 比较金额相等的数据
                if (entry.getKey().compareTo(numRebateInAgreementCurrency) == 0) {
                    // 重命名文件
                    String targetInvoiceID = entry.getValue();
                    String suffixName = StrUtil.subAfter(convertAddAmountFile.getName(), ".", true);
                    boolean renameToFlag = convertAddAmountFile.renameTo(new File(zipInvoicesSourceFileUnZipRenameInvoiceIDPath + File.separator + targetInvoiceID + "." + suffixName));
                    log.info("重命名文件：{}，{}", (zipInvoicesSourceFileUnZipRenameInvoiceIDPath + File.separator + targetInvoiceID + "." + suffixName), renameToFlag ? "成功" : "失败");
                }
            }
        }

    }

    public static List<JobTask1ExcelModelRead> parseContraCogsInvoices(File sourceFile) {
        List<JobTask1ExcelModelRead> contraCogsInvoicesList = new ArrayList<>();
        log.info("解析文件名：{}", sourceFile.getName());
        EasyExcel.read(sourceFile, JobTask1ExcelModelRead.class, new PageReadListener<JobTask1ExcelModelRead>(dataList -> {
            if (CollectionUtils.isNotEmpty(dataList)) {
                contraCogsInvoicesList.addAll(dataList);
            }
        })).sheet(0).doRead();
        return contraCogsInvoicesList;
    }

    public static BigDecimal getOriginalBalanceNum(String value) {
        // $2,301.58
        if (StrUtil.isNotEmpty(value)) {
            value = value.replace("$", "").replace(",", "");
        }

        try {
            value = value.replace(",", "");
            if (value.startsWith("-")) {
                value = value.replace("-", "");
                return new BigDecimal(value).negate();
            }

            return new BigDecimal(value);
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
    }

    /**
     * 合并Invoice ID可用文件
     *
     * @param mergeAvailablePathFile
     * @param availableInvoiceIDFileList
     * @throws Exception
     */
    public void mergeAvailableFile(File mergeAvailablePathFile, List<File> availableInvoiceIDFileList) throws Exception {
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(mergeAvailablePathFile).build();
            WriteSheet writeSheet = EasyExcel.writerSheet(0, "汇总InvoiceID数据")
                    .registerWriteHandler(CustomStyleWriteHandler.buildDefaultWriteHandler(JobTask1ExcelModelWrite3.class))
                    .head(JobTask1ExcelModelWrite3.class).build();

            // 解析目标格式原始文件
            for (File targetFile : availableInvoiceIDFileList) {
                String sourceFileName = targetFile.getName();
                List<JobTask1ExcelModelRead2> targetFileInfoList = parseTargetFile(targetFile);
                if (CollectionUtils.isEmpty(targetFileInfoList)) {
                    continue;
                }
                log.info("原文件名：{}", sourceFileName);

                // 重命名文件
                String targetInvoiceID = StrUtil.subBefore(sourceFileName, ".", false);

                log.info("文件名：{}，匹配金额成功，开始解析", sourceFileName);
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
                log.info("文件名：{}，解析并写入数据完成", sourceFileName);
                log.info("原文件名：{}，解析结束", sourceFileName);
            }
            log.info("所有符合条件的文件解析并写入数据完成");
        } catch (Exception e) {
            throw e;
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }
}