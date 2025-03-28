package com.excel.tools;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.ExcelWriter;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.util.*;
import java.util.stream.Collectors;

public class OnedayExcelGenerator {
    private static final String SAVE_FOLDER = "C:\\excel";

    // 添加进度回调接口
    public interface ProgressCallback {
        void onProgress(int progress, String message);
    }

    // 数据监听器，用于读取CSV文件
    private static class CsvDataListener extends AnalysisEventListener<Map<Integer, String>> {
        private final List<List<String>> rows = new ArrayList<>();
        private List<String> headers;

        @Override
        public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
            headers = new ArrayList<>(headMap.values());
        }

        @Override
        public void invoke(Map<Integer, String> data, AnalysisContext context) {
            List<String> row = new ArrayList<>(data.values());
            rows.add(row);
        }

        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {}

        public List<List<String>> getRows() {
            return rows;
        }

        public List<String> getHeaders() {
            return headers;
        }
    }

    private static String extractDateFromFilename(String filename, String prefix) {
        int startIdx = filename.indexOf(prefix);
        if (startIdx == -1) {
            return null;
        }
        return filename.substring(startIdx + prefix.length(), startIdx + prefix.length() + 8);
    }

    // 修改方法签名，添加进度回调参数
    public static void processFiles(ProgressCallback callback) {
        try {
            callback.onProgress(0, "正在初始化...");
            new File(SAVE_FOLDER).mkdirs();

            JFileChooser mrChooser = new JFileChooser();
            mrChooser.setMultiSelectionEnabled(true);
            mrChooser.setDialogTitle("选择MR文件");
            mrChooser.setFileFilter(new FileNameExtensionFilter("CSV Files", "csv"));

            if (mrChooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
                callback.onProgress(100, "未选择MR文件！");
                return;
            }

            callback.onProgress(10, "选择指标文件...");
            JFileChooser indexChooser = new JFileChooser();
            indexChooser.setMultiSelectionEnabled(true);
            indexChooser.setDialogTitle("选择指标文件");
            indexChooser.setFileFilter(new FileNameExtensionFilter("CSV Files", "csv"));

            if (indexChooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
                callback.onProgress(100, "未选择指标文件！");
                return;
            }

            File[] mrFiles = mrChooser.getSelectedFiles();
            File[] indexFiles = indexChooser.getSelectedFiles();

            callback.onProgress(20, "正在分析文件...");
            Map<String, List<File>> mrGroups = Arrays.stream(mrFiles)
                    .filter(file -> extractDateFromFilename(file.getName(), "MR-") != null)
                    .collect(Collectors.groupingBy(file -> 
                            extractDateFromFilename(file.getName(), "MR-")));

            Map<String, List<File>> indexGroups = Arrays.stream(indexFiles)
                    .filter(file -> extractDateFromFilename(file.getName(), "指标-") != null)
                    .collect(Collectors.groupingBy(file -> 
                            extractDateFromFilename(file.getName(), "指标-")));

            Set<String> commonDates = new HashSet<>(mrGroups.keySet());
            commonDates.retainAll(indexGroups.keySet());

            if (commonDates.isEmpty()) {
                callback.onProgress(100, "未找到MR文件和指标文件的共同日期！");
                return;
            }

            int totalDates = commonDates.size();
            int currentDate = 0;

            for (String date : commonDates) {
                callback.onProgress(30 + (currentDate * 60 / totalDates), 
                    "正在处理日期: " + date + " (" + (currentDate + 1) + "/" + totalDates + ")");

                List<List<String>> mrData = new ArrayList<>();
                List<String> mrHeaders = null;
                for (File file : mrGroups.get(date)) {
                    CsvDataListener listener = new CsvDataListener();
                    EasyExcel.read(file.getAbsolutePath())
                            .charset(java.nio.charset.Charset.forName("GBK"))
                            .sheet()
                            .registerReadListener(listener)
                            .doRead();
                    if (mrHeaders == null) {
                        mrHeaders = listener.getHeaders();
                    }
                    mrData.addAll(listener.getRows());
                }

                List<List<String>> indexData = new ArrayList<>();
                List<String> indexHeaders = null;
                for (File file : indexGroups.get(date)) {
                    CsvDataListener listener = new CsvDataListener();
                    EasyExcel.read(file.getAbsolutePath())
                            .charset(java.nio.charset.Charset.forName("GBK"))
                            .sheet()
                            .registerReadListener(listener)
                            .doRead();
                    if (indexHeaders == null) {
                        indexHeaders = listener.getHeaders();
                    }
                    indexData.addAll(listener.getRows());
                }

                String savePath = SAVE_FOLDER + "\\" + date + ".xlsx";
                
                try (ExcelWriter excelWriter = EasyExcel.write(savePath).build()) {
                    WriteSheet mrSheet = EasyExcel.writerSheet(0, "MR")
                            .head(convertToHead(mrHeaders))
                            .build();
                    excelWriter.write(formatDataList(mrData), mrSheet);

                    WriteSheet indexSheet = EasyExcel.writerSheet(1, "指标")
                            .head(convertToHead(indexHeaders))
                            .build();
                    excelWriter.write(formatDataList(indexData), indexSheet);
                }

                callback.onProgress(90 + (currentDate * 10 / totalDates), 
                    "正在保存文件: " + date + ".xlsx");
                currentDate++;
            }

            callback.onProgress(100, "处理完成");

        } catch (Exception e) {
            e.printStackTrace();
            callback.onProgress(100, "处理出错: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                processFiles((progress, message) -> 
                    System.out.println("Progress: " + progress + "%, Message: " + message));
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

    // 修改数据格式转换方法，保持原始数值格式
    private static List<List<Object>> formatDataList(List<List<String>> data) {
        return data.stream()
                .map(row -> row.stream()
                        .map(item -> {
                            if (item == null || item.trim().isEmpty()) {
                                return (Object)"";
                            }
                            try {
                                // 尝试将字符串转换为数值，保持原始格式
                                double value = Double.parseDouble(item.trim());
                                return (Object)value;  // 直接返回数值，不做格式化
                            } catch (NumberFormatException e) {
                                // 如果不是数值，则直接返回原字符串
                                return (Object)item;
                            }
                        })
                        .collect(Collectors.toList()))
                .collect(Collectors.toCollection(ArrayList::new));
    }

    // 表头转换方法保持不变
    private static List<List<String>> convertToHead(List<String> headers) {
        return headers.stream()
                .map(Collections::singletonList)
                .collect(Collectors.toList());
    }
}