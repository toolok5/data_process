package com.excel.tools;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;

public class ExcelToolsLauncher {
    private static JProgressBar progressBar;
    private static Timer progressTimer;
    private static JFrame mainFrame;  // 添加静态引用以便控制主窗口
    
    static {
        // 禁用SLF4J警告
        System.setProperty("org.slf4j.simpleLogger.defaultLogLevel", "off");
    }
    
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            mainFrame = new JFrame("Excel工具集");
            mainFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
            mainFrame.setSize(400, 350);
            mainFrame.setLocationRelativeTo(null);

            JPanel mainPanel = new JPanel(new BorderLayout());
            mainPanel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

            // 创建进度条面板
            JPanel progressPanel = new JPanel(new BorderLayout());
            progressPanel.setBorder(BorderFactory.createEmptyBorder(0, 0, 10, 0));
            
            progressBar = new JProgressBar(0, 100);
            progressBar.setStringPainted(true);
            progressBar.setString("就绪");
            progressBar.setValue(0);
            progressPanel.add(progressBar, BorderLayout.CENTER);
            
            mainPanel.add(progressPanel, BorderLayout.NORTH);

            JPanel buttonPanel = new JPanel();
            buttonPanel.setLayout(new GridLayout(3, 1, 10, 10));

            JButton btnRowCounter = new JButton("Excel行数统计工具");
            JButton btnDailyGenerator = new JButton("每日Excel生成工具");
            JButton btnExcelGenerator = new JButton("Excel生成工具");

            btnRowCounter.addActionListener((ActionEvent e) -> {
                launchToolInNewThread(() -> 
                    ExcelRowCounter.countFileRows((progress, message) -> {
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(progress);
                            progressBar.setString(message);
                        });
                    }), 
                "Excel行数统计工具");
            });

            btnDailyGenerator.addActionListener((ActionEvent e) -> {
                launchToolInNewThread(() -> 
                    OnedayExcelGenerator.processFiles((progress, message) -> {
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(progress);
                            progressBar.setString(message);
                        });
                    }), 
                "每日Excel生成工具");
            });

            btnExcelGenerator.addActionListener((ActionEvent e) -> {
                launchToolInNewThread(() -> 
                    ExcelGenerator.processFiles((progress, message) -> {
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(progress);
                            progressBar.setString(message);
                        });
                    }), 
                "Excel生成工具");
            });

            buttonPanel.add(btnRowCounter);
            buttonPanel.add(btnDailyGenerator);
            buttonPanel.add(btnExcelGenerator);

            mainPanel.add(buttonPanel, BorderLayout.CENTER);
            mainFrame.add(mainPanel);
            mainFrame.setVisible(true);
        });
    }

    private static void launchToolInNewThread(Runnable tool, String toolName) {
        // 禁用所有按钮，防止重复点击
        disableButtons();
        
        // 重置并启动进度条
        progressBar.setValue(0);
        progressBar.setString(toolName + " 运行中...");
        
        if (progressTimer != null && progressTimer.isRunning()) {
            progressTimer.stop();
        }
        
        progressTimer = new Timer(50, e -> {
            int value = progressBar.getValue();
            if (value < 90) {
                progressBar.setValue(value + 1);
            }
        });
        progressTimer.start();

        // 使用自定义ThreadGroup来防止System.exit()
        ThreadGroup noExitGroup = new ThreadGroup("NoExit") {
            @Override
            public void uncaughtException(Thread t, Throwable e) {
                if (!(e instanceof SecurityException)) {
                    super.uncaughtException(t, e);
                }
            }
        };

        Thread toolThread = new Thread(noExitGroup, () -> {
            try {
                tool.run();
                SwingUtilities.invokeLater(() -> {
                    progressTimer.stop();
                    progressBar.setValue(100);
                    progressBar.setString(toolName + " 已完成");
                    
                    Timer resetTimer = new Timer(1000, evt -> {
                        progressBar.setValue(0);
                        progressBar.setString("就绪");
                        ((Timer)evt.getSource()).stop();
                        enableButtons();
                    });
                    resetTimer.setRepeats(false);
                    resetTimer.start();
                });
            } catch (Exception e) {
                SwingUtilities.invokeLater(() -> {
                    progressTimer.stop();
                    progressBar.setValue(0);
                    progressBar.setString(toolName + " 运行出错: " + e.getMessage());
                    enableButtons();
                });
            }
        });
        
        toolThread.setDaemon(true);
        toolThread.start();
    }
    
    private static void disableButtons() {
        Container contentPane = mainFrame.getContentPane();
        if (contentPane.getComponent(0) instanceof JPanel) {
            JPanel mainPanel = (JPanel) contentPane.getComponent(0);
            if (mainPanel.getComponent(1) instanceof JPanel) {
                JPanel buttonPanel = (JPanel) mainPanel.getComponent(1);
                for (Component comp : buttonPanel.getComponents()) {
                    if (comp instanceof JButton) {
                        comp.setEnabled(false);
                    }
                }
            }
        }
    }
    
    private static void enableButtons() {
        Container contentPane = mainFrame.getContentPane();
        if (contentPane.getComponent(0) instanceof JPanel) {
            JPanel mainPanel = (JPanel) contentPane.getComponent(0);
            if (mainPanel.getComponent(1) instanceof JPanel) {
                JPanel buttonPanel = (JPanel) mainPanel.getComponent(1);
                for (Component comp : buttonPanel.getComponents()) {
                    if (comp instanceof JButton) {
                        comp.setEnabled(true);
                    }
                }
            }
        }
    }
}