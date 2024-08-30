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
import javafx.scene.control.Separator;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.openjfx.easyexcel.CustomStyleWriteHandler;
import org.openjfx.utils.file.FileUtil;
import org.openjfx.model.JobTask1ExcelModelRead;
import org.openjfx.model.JobTask1ExcelModelRead2;
import org.openjfx.model.JobTask1ExcelModelRead3;
import org.openjfx.model.JobTask1ExcelModelWrite3;

import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 匹配发票并合并发票v2
 */
@Slf4j
public class InvoiceReportFileHandlerAppV2 extends Application {

    private Label fileLabel1;
    private Label fileLabel2;
    private File contraCogsInvoicesSourceFile;
    private File zipInvoicesSourceFile;

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("发票处理");
        primaryStage.getIcons().add(new Image("file:icon.png")); // 需要将 icon.png 放在项目根目录

        // 添加“按钮1”标签
        Label button1Label = new Label("ContraCogsInvoices");
        Button uploadButton1 = new Button("上传");
        uploadButton1.setGraphic(new ImageView(new Image("file:upload_icon.png"))); // 需要将 upload_icon.png 放在项目根目录
        fileLabel1 = new Label("请选择文件");
        fileLabel1.setTextFill(Color.GRAY);

        uploadButton1.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(primaryStage);
            if (file != null) {
                contraCogsInvoicesSourceFile = file;
                fileLabel1.setText("文件名：" + contraCogsInvoicesSourceFile.getName());
                fileLabel1.setTextFill(Color.BLACK);
            } else {
                fileLabel1.setText("请选择文件");
                fileLabel1.setTextFill(Color.RED);  // 若未选择文件则变为红色提示
            }
        });

        // 添加“按钮2”标签
        Label button2Label = new Label("待处理发票压缩包");
        Button uploadButton2 = new Button("上传");
        uploadButton2.setGraphic(new ImageView(new Image("file:upload_icon.png")));
        fileLabel2 = new Label("请选择文件");
        fileLabel2.setTextFill(Color.GRAY);

        uploadButton2.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(primaryStage);
            if (file != null) {
                zipInvoicesSourceFile = file;
                fileLabel2.setText("文件名：" + zipInvoicesSourceFile.getName());
                fileLabel2.setTextFill(Color.BLACK);
            } else {
                fileLabel2.setText("请选择文件");
                fileLabel2.setTextFill(Color.RED);  // 若未选择文件则变为红色提示
            }
        });

        // 提交按钮
        Button submitButton = new Button("提交");
        submitButton.setStyle("-fx-background-color: #4CAF50; -fx-text-fill: white; -fx-font-size: 14px;");
        submitButton.setOnAction(e -> {
            boolean valid = true;

            if (contraCogsInvoicesSourceFile == null) {
                fileLabel1.setText("此项必填");
                fileLabel1.setTextFill(Color.RED);
                valid = false;
            }

            if (zipInvoicesSourceFile == null) {
                fileLabel2.setText("此项必填");
                fileLabel2.setTextFill(Color.RED);
                valid = false;
            }

            String zipInvoicesSourceFileName = zipInvoicesSourceFile.getName();
            String zipInvoicesSourceFileSuffixName = StrUtil.subAfter(zipInvoicesSourceFileName, ".", true);
            if (!"zip".equalsIgnoreCase(zipInvoicesSourceFileSuffixName)) {
                fileLabel2.setText("仅限上传zip压缩包");
                fileLabel2.setTextFill(Color.RED);
                valid = false;
            }

            if (!valid) {
                return;
            }

            submitButton.setDisable(true);
            submitButton.setText("处理中...");

            Task<Void> task = new Task<Void>() {
                @Override
                protected Void call() throws Exception {
                    processTask(contraCogsInvoicesSourceFile, zipInvoicesSourceFile);
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
        });

        // 重置按钮
        Button resetButton = new Button("重置");
        resetButton.setStyle("-fx-background-color: #f44336; -fx-text-fill: white; -fx-font-size: 14px;");
        resetButton.setOnAction(e -> {
            contraCogsInvoicesSourceFile = null;
            zipInvoicesSourceFile = null;
            fileLabel1.setText("请选择文件");
            fileLabel1.setTextFill(Color.GRAY);
            fileLabel2.setText("请选择文件");
            fileLabel2.setTextFill(Color.GRAY);
        });

        // 布局管理器
        GridPane gridPane = new GridPane();
        gridPane.setPadding(new Insets(20));
        gridPane.setVgap(10);
        gridPane.setHgap(10);
        gridPane.setAlignment(Pos.CENTER);

        gridPane.add(button1Label, 0, 0);    // 在第0行，第0列添加“按钮1”标签
        gridPane.add(uploadButton1, 1, 0);   // 在第0行，第1列添加上传按钮1
        gridPane.add(fileLabel1, 2, 0);      // 在第0行，第2列添加文件标签1

        gridPane.add(button2Label, 0, 1);    // 在第1行，第0列添加“按钮2”标签
        gridPane.add(uploadButton2, 1, 1);   // 在第1行，第1列添加上传按钮2
        gridPane.add(fileLabel2, 2, 1);      // 在第1行，第2列添加文件标签2

        // 将提交按钮和重置按钮放在一起
        HBox actionBox = new HBox(10, submitButton, resetButton);
        actionBox.setAlignment(Pos.CENTER);

        Separator separator = new Separator();
        separator.setPadding(new Insets(10, 0, 10, 0));

        VBox vbox = new VBox(gridPane, separator, actionBox);
        vbox.setSpacing(10);

        // 调整窗口尺寸，宽度550，高度275
        Scene scene = new Scene(vbox, 700, 400);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void showAlert(Alert.AlertType alertType, String title, String message) {
        Alert alert = new Alert(alertType);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

    public void processTask(File contraCogsInvoicesSourceFile, File zipInvoicesSourceFile) throws Exception {
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

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
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
