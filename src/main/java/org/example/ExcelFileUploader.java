package org.example;

import javax.swing.*;
import java.awt.Desktop;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelFileUploader extends JFrame {

    private JButton btnSelectSource;
    private JButton btnSelectConfig;
    private JButton btnProcess;
    private JButton btnOpenResult; // ✅ Thêm nút mở file
    private JLabel statusLabel;
    private JLabel loadingLabel;
    private final FileProcesserService fileProcesserService;

    private String sourcePath;
    private String configPath;
    private String targetPath;

    public ExcelFileUploader(FileProcesserService fileProcesserService) {
        this.fileProcesserService = fileProcesserService;

        setTitle("Excel File Processor");
        setSize(500, 350); // Tăng chiều cao để đủ chỗ cho nút mới
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(null);

        // Nút chọn file nguồn
        btnSelectSource = new JButton("Chọn file nguồn");
        btnSelectSource.setBounds(50, 30, 180, 30);
        add(btnSelectSource);

        // Nút chọn file cấu hình
        btnSelectConfig = new JButton("Chọn file cấu hình");
        btnSelectConfig.setBounds(260, 30, 180, 30);
        add(btnSelectConfig);

        // Nút xử lý
        btnProcess = new JButton("Xử lý file");
        btnProcess.setBounds(150, 80, 180, 30);
        btnProcess.setEnabled(false);
        add(btnProcess);

        // Label trạng thái
        statusLabel = new JLabel("Vui lòng chọn file nguồn và file cấu hình.");
        statusLabel.setBounds(50, 130, 400, 30);
        add(statusLabel);

        // Loading
        loadingLabel = new JLabel();
        loadingLabel.setBounds(200, 170, 100, 50);
        loadingLabel.setVisible(false);
        loadingLabel.setIcon(new ImageIcon(getClass().getResource("/loading.gif")));
        add(loadingLabel);

        // ✅ Nút mở file kết quả
        btnOpenResult = new JButton("Mở file kết quả");
        btnOpenResult.setBounds(150, 230, 180, 30);
        btnOpenResult.setEnabled(false);
        add(btnOpenResult);

        // Chọn file nguồn
        btnSelectSource.addActionListener((ActionEvent e) -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Excel files", "xls", "xlsx"));
            int result = fileChooser.showOpenDialog(ExcelFileUploader.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                sourcePath = selectedFile.getAbsolutePath();

                // Tự động tạo targetPath cùng vị trí, thêm hậu tố _result
                String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
                int dotIndex = sourcePath.lastIndexOf('.');
                if (dotIndex != -1) {
                    targetPath = sourcePath.substring(0, dotIndex)
                            + "_result_" + timestamp
                            + sourcePath.substring(dotIndex);
                } else {
                    targetPath = sourcePath + "_result_" + timestamp + ".xlsx";
                }

                updateStatus();
            }
        });

        // Chọn file cấu hình
        btnSelectConfig.addActionListener((ActionEvent e) -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Chọn file cấu hình");
            fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Excel files", "xls", "xlsx"));
            int result = fileChooser.showOpenDialog(ExcelFileUploader.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                configPath = fileChooser.getSelectedFile().getAbsolutePath();
                updateStatus();
            }
        });

        // Xử lý file
        btnProcess.addActionListener((ActionEvent e) -> {
            loadingLabel.setVisible(true);
            btnProcess.setEnabled(false);
            btnOpenResult.setEnabled(false); // ✅ disable trước khi xử lý

            new Thread(() -> {
                try {
                    fileProcesserService.processing(sourcePath, targetPath, configPath);

                    SwingUtilities.invokeLater(() -> {
                        loadingLabel.setVisible(false);
                        btnProcess.setEnabled(true);
                        btnOpenResult.setEnabled(true); // ✅ enable sau khi xử lý
                        statusLabel.setText("✅ Xử lý hoàn tất.");
                        JOptionPane.showMessageDialog(
                                ExcelFileUploader.this,
                                "✅ File đã xử lý xong:\n" + targetPath,
                                "Thông báo",
                                JOptionPane.INFORMATION_MESSAGE
                        );
                    });
                } catch (IOException ex) {
                    ex.printStackTrace();
                    SwingUtilities.invokeLater(() -> {
                        loadingLabel.setVisible(false);
                        btnProcess.setEnabled(true);
                        statusLabel.setText("❌ Lỗi xử lý file.");
                        JOptionPane.showMessageDialog(
                                ExcelFileUploader.this,
                                "❌ Lỗi khi xử lý file:\n" + ex.getMessage(),
                                "Lỗi",
                                JOptionPane.ERROR_MESSAGE
                        );
                    });
                }
            }).start();
        });

        // ✅ Sự kiện mở file kết quả
        btnOpenResult.addActionListener((ActionEvent e) -> {
            if (targetPath != null) {
                try {
                    Desktop.getDesktop().open(new File(targetPath));
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(
                            ExcelFileUploader.this,
                            "❌ Không thể mở file:\n" + ex.getMessage(),
                            "Lỗi",
                            JOptionPane.ERROR_MESSAGE
                    );
                }
            }
        });
    }

    private void updateStatus() {
        if (sourcePath != null && configPath != null) {
            statusLabel.setText("✅ Đã chọn đủ file. Bấm 'Xử lý file' để tiếp tục.");
            btnProcess.setEnabled(true);
        } else if (sourcePath != null) {
            statusLabel.setText("🔹 Đã chọn file nguồn. Chọn thêm file cấu hình.");
        } else if (configPath != null) {
            statusLabel.setText("🔹 Đã chọn file cấu hình. Chọn thêm file nguồn.");
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            FileProcesserService service = new FileProcesserService();
            ExcelFileUploader uploader = new ExcelFileUploader(service);
            uploader.setVisible(true);
        });
    }
}
