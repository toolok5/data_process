package com.excel.tools;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.ExcelReader;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelRowCounter {
    // 添加进度回调接口
    public interface ProgressCallback {
        void onProgress(int progress, String message);
    }
    
    private static class RowCountListener extends AnalysisEventListener<Map<Integer, String>> {
        private int rowCount = 0;
        private boolean isFirstRow = true;
        private final ProgressCallback callback;
        private final int totalFiles;
        private final int currentFileIndex;
        private int currentRow = 0;
        private static final int ESTIMATED_ROWS = 1000; // 估计的行数
        
        public RowCountListener(ProgressCallback callback, int totalFiles, int currentFileIndex) {
            this.callback = callback;
            this.totalFiles = totalFiles;
            this.currentFileIndex = currentFileIndex;
        }

        @Override
        public void invoke(Map<Integer, String> data, AnalysisContext context) {
            boolean isRowEmpty = data.values().stream()
                    .allMatch(value -> value == null || value.trim().isEmpty());
            
            if (!isRowEmpty && !isFirstRow) {
                rowCount++;
                currentRow = context.readRowHolder().getRowIndex();
                // 使用估计的总行数计算进度
                int fileProgress = Math.min(currentRow * 100 / ESTIMATED_ROWS, 100);
                int totalProgress = (currentFileIndex * 100 / totalFiles) + 
                                  (fileProgress / totalFiles);
                callback.onProgress(Math.min(totalProgress, 90), 
                    "正在处理第 " + (currentFileIndex + 1) + "/" + totalFiles + " 个文件");
            }
            isFirstRow = false;
        }

        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {}

        public int getRowCount() {
            return rowCount;
        }
    }

    private static class FileStats {
        String fileName;
        String sheetName;
        int rowCount;

        public FileStats(String fileName, String sheetName, int rowCount) {
            this.fileName = fileName;
            this.sheetName = sheetName;
            this.rowCount = rowCount;
        }

        public Map<String, Object> toMap() {
            Map<String, Object> map = new HashMap<>();
            map.put("文件名", fileName);
            map.put("工作表名", sheetName);
            map.put("行数", rowCount);
            return map;
        }
    }

    private static List<Integer> parseColumnRange(String columnRange) {
        try {
            List<Integer> columns = new ArrayList<>();
            if (columnRange.contains("-")) {
                String[] range = columnRange.split("-");
                int start = Integer.parseInt(range[0].trim()) - 1;
                int end = Integer.parseInt(range[1].trim());
                for (int i = start; i < end; i++) {
                    columns.add(i);
                }
            } else {
                columns.add(Integer.parseInt(columnRange.trim()) - 1);
            }
            return columns;
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("无效的列范围输入，请输入数字或范围（例如：1 或 1-3）");
        }
    }

    private static List<FileStats> getRowCount(String filePath, List<Integer> columns, 
            ProgressCallback callback, int totalFiles, int currentFileIndex) {
        List<FileStats> results = new ArrayList<>();
        String fileName = new File(filePath).getName();

        try {
            if (filePath.endsWith(".csv")) {
                RowCountListener listener = new RowCountListener(callback, totalFiles, currentFileIndex);
                EasyExcel.read(filePath)
                        .sheet()
                        .headRowNumber(0) // CSV文件不跳过表头行
                        .registerReadListener(listener)
                        .doRead();
                results.add(new FileStats(fileName, "CSV", listener.getRowCount()));
            } else if (filePath.endsWith(".xlsx") || filePath.endsWith(".xls")) {
                ExcelReader workbook = EasyExcel.read(filePath).build();
                List<ReadSheet> sheets = workbook.excelExecutor().sheetList();
                
                for (ReadSheet sheet : sheets) {
                    String sheetName = sheet.getSheetName();
                    RowCountListener listener = new RowCountListener(callback, totalFiles, currentFileIndex);
                    EasyExcel.read(filePath)
                            .sheet(sheetName)
                            .headRowNumber(1) // Excel文件跳过表头行
                            .registerReadListener(listener)
                            .doRead();
                    results.add(new FileStats(fileName, sheetName, listener.getRowCount()));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            results.add(new FileStats(fileName, "错误", -1));
        }
        return results;
    }

    private static void appendToExistingFile(String existingPath, List<Map<String, Object>> newData) {
        try {
            // 读取现有数据
            List<Map<Integer, Object>> existingRows = new ArrayList<>();
            int maxRowCount = 0;
            
            if (Files.exists(Paths.get(existingPath))) {
                // 读取现有数据并找出最大行数
                EasyExcel.read(existingPath)
                        .sheet()
                        .registerReadListener(new AnalysisEventListener<Map<Integer, Object>>() {
                            @Override
                            public void invoke(Map<Integer, Object> data, AnalysisContext context) {
                                if (context.readRowHolder().getRowIndex() > 0) {
                                    existingRows.add(data);
                                }
                            }

                            @Override
                            public void doAfterAllAnalysed(AnalysisContext context) {}
                        })
                        .doRead();
                
                maxRowCount = existingRows.size();
            }

            // 确定新列的起始索引（每组3列：文件名、工作表名、行数）
            int groupCount = existingRows.isEmpty() ? 0 : (existingRows.get(0).size() / 3);
            
            // 准备合并后的数据
            List<List<Object>> mergedData = new ArrayList<>();
            
            // 添加表头
            List<String> headers = new ArrayList<>();
            // 添加现有表头
            for (int i = 0; i < groupCount; i++) {
                headers.add("文件名_" + (i + 1));
                headers.add("工作表名_" + (i + 1));
                headers.add("行数_" + (i + 1));
            }
            // 添加新数据的表头
            headers.add("文件名_" + (groupCount + 1));
            headers.add("工作表名_" + (groupCount + 1));
            headers.add("行数_" + (groupCount + 1));
            
            List<List<String>> headList = headers.stream()
                    .map(Collections::singletonList)
                    .collect(Collectors.toList());
            
            // 准备所有行数据
            int totalRows = Math.max(maxRowCount, newData.size());
            for (int i = 0; i < totalRows; i++) {
                List<Object> rowData = new ArrayList<>();
                
                // 添加现有数据
                if (i < existingRows.size()) {
                    Map<Integer, Object> existingRow = existingRows.get(i);
                    for (int j = 0; j < existingRow.size(); j++) {
                        rowData.add(String.valueOf(existingRow.getOrDefault(j, "")));
                    }
                } else {
                    // 填充空值
                    for (int j = 0; j < groupCount * 3; j++) {
                        rowData.add("");
                    }
                }
                
                // 添加新数据
                if (i < newData.size()) {
                    Map<String, Object> newRow = newData.get(i);
                    rowData.add(newRow.get("文件名"));
                    rowData.add(newRow.get("工作表名"));
                    rowData.add(newRow.get("行数"));
                } else {
                    // 填充空值
                    rowData.add("");
                    rowData.add("");
                    rowData.add("");
                }
                
                mergedData.add(rowData);
            }

            // 写入合并后的数据
            EasyExcel.write(existingPath)
                    .head(headList)
                    .sheet("统计结果")
                    .doWrite(mergedData);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void countFileRows(ProgressCallback progressCallback) {
        try {
            String columnRange = JOptionPane.showInputDialog(
                    null,
                    "请输入需要读取的列范围（默认: 1，例如: 1 或 1-3）：",
                    "1"
            );

            if (columnRange == null || columnRange.trim().isEmpty()) {
                progressCallback.onProgress(100, "操作已取消");
                return;
            }

            List<Integer> columns = parseColumnRange(columnRange);
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setMultiSelectionEnabled(true);
            fileChooser.setFileFilter(new FileNameExtensionFilter(
                    "Excel Files", "xlsx", "xls", "csv"));

            if (fileChooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
                progressCallback.onProgress(100, "未选择任何文件");
                return;
            }

            File[] files = fileChooser.getSelectedFiles();
            List<Map<String, Object>> stats = new ArrayList<>();

            for (int i = 0; i < files.length; i++) {
                File file = files[i];
                progressCallback.onProgress(i * 90 / files.length, 
                    "正在处理第 " + (i + 1) + "/" + files.length + " 个文件: " + file.getName());
                
                List<FileStats> fileStats = getRowCount(file.getAbsolutePath(), columns, 
                    progressCallback, files.length, i);
                
                for (FileStats stat : fileStats) {
                    stats.add(stat.toMap());
                }
            }

            progressCallback.onProgress(95, "正在保存结果...");
            
            // 保存结果
            String saveDir = "C:\\excel";
            new File(saveDir).mkdirs();
            String savePath = saveDir + "\\行数统计结果.xlsx";

            if (new File(savePath).exists()) {
                appendToExistingFile(savePath, stats);
            } else {
                List<List<Object>> rows = stats.stream()
                    .map(stat -> Arrays.asList(
                        stat.get("文件名"),
                        stat.get("工作表名"),
                        stat.get("行数")))
                    .collect(Collectors.toList());

                List<List<String>> headList = Arrays.asList(
                    Collections.singletonList("文件名"),
                    Collections.singletonList("工作表名"),
                    Collections.singletonList("行数")
                );

                EasyExcel.write(savePath)
                        .head(headList)
                        .sheet("统计结果")
                        .doWrite(rows);
            }
            
            progressCallback.onProgress(100, "处理完成");
            
        } catch (Exception e) {
            e.printStackTrace();
            progressCallback.onProgress(100, "处理出错: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                countFileRows((progress, message) -> 
                    System.out.println("Progress: " + progress + "%, Message: " + message));
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }
} 