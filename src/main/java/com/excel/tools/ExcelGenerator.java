package com.excel.tools;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.ExcelWriter;
import javax.swing.*;
import java.io.File;
// import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.io.BufferedReader;
import java.io.IOException;
import java.util.Arrays;
import java.io.InputStreamReader;
import java.io.FileInputStream;

public class ExcelGenerator {
    private static final String FOLDER_PATH = "C:\\excel";

    public interface ProgressCallback {
        void onProgress(int progress, String message);
    }

    public static void processFiles(ProgressCallback callback) {
        try {
            callback.onProgress(0, "正在初始化...");

            String nRangeStr = JOptionPane.showInputDialog(
                "请输入提取到第一个表的列数据范围，用逗号分隔（例如 1-3,5-9):"
            );
            
            if (nRangeStr == null || nRangeStr.trim().isEmpty()) {
                callback.onProgress(100, "操作已取消");
                return;
            }

            String mRangeStr = JOptionPane.showInputDialog(
                "请输入提取到第二个表的列数据范围，用逗号分隔（例如 1-5,7-9):"
            );

            if (mRangeStr == null || mRangeStr.trim().isEmpty()) {
                callback.onProgress(100, "操作已取消");
                return;
            }

            callback.onProgress(10, "正在解析列范围...");
            List<Integer> nColumnsRange = parseColumnRange(nRangeStr);
            List<Integer> mColumnsRange = parseColumnRange(mRangeStr);

            callback.onProgress(20, "正在扫描文件...");
            File folder = new File(FOLDER_PATH);
            File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".csv"));
            
            if (files == null || files.length == 0) {
                callback.onProgress(100, "未找到CSV文件");
                return;
            }

            int totalFiles = files.length;
            for (int i = 0; i < files.length; i++) {
                File file = files[i];
                callback.onProgress(30 + (i * 60 / totalFiles), 
                    "正在处理文件 " + (i + 1) + "/" + totalFiles + ": " + file.getName());
                
                processCsvFile(file, nColumnsRange, mColumnsRange);
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

    private static List<Integer> parseColumnRange(String rangeStr) {
        // 预估容量，避免动态扩容
        List<Integer> columnIndices = new ArrayList<>(32);
        String[] parts = rangeStr.split(",");
        
        for (String part : parts) {
            if (part.contains("-")) {
                String[] range = part.split("-");
                int start = Integer.parseInt(range[0].trim()) - 1;
                int end = Integer.parseInt(range[1].trim());
                // 使用预计算的容量
                int capacity = end - start;
                for (int i = start; i < end; i++) {
                    columnIndices.add(i);
                }
            } else {
                columnIndices.add(Integer.parseInt(part.trim()) - 1);
            }
        }
        return columnIndices;
    }

    private static void processCsvFile(File csvFile, List<Integer> nColumnsRange, List<Integer> mColumnsRange) {
        try {
            // 读取CSV文件
            List<List<String>> allData = CsvReader.readCsv(csvFile.getPath());
            
            // 提取指定列的数据
            List<List<String>> originalData = extractColumns(allData, nColumnsRange);
            List<List<String>> processedData = extractColumns(allData, mColumnsRange);

            // 生成输出Excel文件路径
            String outputPath = csvFile.getPath().replace(".csv", ".xlsx");

            // 写入Excel文件
            writeToExcel(outputPath, originalData, processedData);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<List<String>> extractColumns(List<List<String>> data, List<Integer> columnIndices) {
        List<List<String>> result = new ArrayList<>();
        for (List<String> row : data) {
            List<String> newRow = new ArrayList<>();
            for (Integer index : columnIndices) {
                if (index < row.size()) {
                    newRow.add(row.get(index));
                }
            }
            result.add(newRow);
        }
        return result;
    }

    private static void writeToExcel(String outputPath, 
                             List<List<String>> originalData, 
                             List<List<String>> processedData) {
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(outputPath)
                    .charset(java.nio.charset.Charset.forName("GBK"))
                    .build();

            // 写入第一个工作表
            WriteSheet writeSheet1 = EasyExcel.writerSheet(0, "KPI,MR处理").build();
            excelWriter.write(originalData, writeSheet1);

            // 写入第二个工作表
            WriteSheet writeSheet2 = EasyExcel.writerSheet(1, "数据处理").build();
            excelWriter.write(processedData, writeSheet2);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // 确保ExcelWriter被正确关闭
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }
}

// CSV读取器类
class CsvReader {
    private static final int BUFFER_SIZE = 8192;  // 设置更大的缓冲区
    
    public static List<List<String>> readCsv(String filePath) {
        List<List<String>> records = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(new FileInputStream(filePath), "GBK"), 
                BUFFER_SIZE)) {  // 使用更大的缓冲区
            String line;
            while ((line = br.readLine()) != null) {
                // 预分配容量，避免动态扩容
                List<String> values = new ArrayList<>(Arrays.asList(line.split(",")));
                records.add(values);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return records;
    }
}