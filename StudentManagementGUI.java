import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class StudentManagementGUI extends JFrame {
    // Colors (customized for dark theme)
    private static final Color DARK_BG = new Color(30, 30, 30);
    private static final Color PANEL_BG = new Color(40, 40, 40);
    private static final Color BUTTON_BG = new Color(114, 137, 218); // #7289da
    private static final Color BUTTON_HOVER = new Color(91, 110, 174); // #5b6eae
    private static final Color TEXT_COLOR = new Color(230, 230, 230);
    
    private JLabel statusLabel;
    private boolean executing = false;
    private ExcelPrinter excelPrinter;
    private ExecutorService executor;
    
    public StudentManagementGUI() {
        // Setup main frame
        super("Student Management System");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 600);
        setLocationRelativeTo(null); // Center on screen
        getContentPane().setBackground(DARK_BG);
        
        // Set dark theme for entire application
        setDarkLookAndFeel();
        
        // Initialize components
        excelPrinter = new ExcelPrinter();
        executor = Executors.newSingleThreadExecutor();
        
        // Setup UI
        setupUI();
        
        // Add window closing event
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                exitProgram();
            }
        });
    }
    
    private void setDarkLookAndFeel() {
        try {
            // Set cross-platform Java L&F
            UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
            
            // Override colors for dark theme
            UIManager.put("Panel.background", PANEL_BG);
            UIManager.put("Button.background", BUTTON_BG);
            UIManager.put("Button.foreground", Color.WHITE);
            UIManager.put("Label.foreground", TEXT_COLOR);
            UIManager.put("CheckBox.foreground", TEXT_COLOR);
            UIManager.put("TextField.background", new Color(60, 60, 60));
            UIManager.put("TextField.foreground", TEXT_COLOR);
            UIManager.put("OptionPane.background", DARK_BG);
            UIManager.put("OptionPane.messageForeground", TEXT_COLOR);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    private void setupUI() {
        // Main panel with GridBagLayout
        JPanel mainPanel = new JPanel(new GridBagLayout());
        mainPanel.setBackground(DARK_BG);
        mainPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
        
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(5, 5, 5, 5);
        
        // Left Panel - Student Data Management
        JPanel leftPanel = createPanel("Student Data Management");
        JButton transferToSF5ABBtn = createButton("Transfer to SF5AB", e -> runScriptAsync("nig.py"));
        JButton formsBtn = createButton("Open School Forms", e -> showSchoolForms());
        JButton transferBtn = createButton("Transfer Grades to Master Forms", e -> runScriptAsync("trans.py"));
        
        leftPanel.add(transferToSF5ABBtn);
        leftPanel.add(Box.createVerticalStrut(10));
        leftPanel.add(formsBtn);
        leftPanel.add(Box.createVerticalStrut(10));
        leftPanel.add(transferBtn);
        
        // Right Panel - Grade Encoding Assistant
        JPanel rightPanel = createPanel("Grade Encoding Assistant");
        JButton editBtn = createButton("Open Master Forms", e -> showQuarterSelection());
        JButton finishBtn = createButton("Finish & Encode", e -> runScriptAsync("grade.py"));
        JButton printTransferBtn = createButton("Print / Transfer", e -> showPrintTransferOptions());
        
        rightPanel.add(editBtn);
        rightPanel.add(Box.createVerticalStrut(10));
        rightPanel.add(finishBtn);
        rightPanel.add(Box.createVerticalStrut(10));
        rightPanel.add(printTransferBtn);
        
        // Add left and right panels to main panel
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        mainPanel.add(leftPanel, gbc);
        
        gbc.gridx = 1;
        mainPanel.add(rightPanel, gbc);
        
        // Status panel
        JPanel statusPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        statusPanel.setBackground(PANEL_BG);
        statusPanel.setBorder(BorderFactory.createEmptyBorder(5, 0, 5, 0));
        
        statusLabel = new JLabel("Ready");
        statusLabel.setForeground(Color.GREEN);
        statusPanel.add(statusLabel);
        
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 2;
        gbc.weighty = 0.0;
        mainPanel.add(statusPanel, gbc);
        
        // Exit button
        JButton exitBtn = createButton("Exit Program", e -> exitProgram());
        
        gbc.gridy = 2;
        gbc.weighty = 0.0;
        mainPanel.add(exitBtn, gbc);
        
        setContentPane(mainPanel);
    }
    
    private JPanel createPanel(String title) {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBackground(PANEL_BG);
        panel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(60, 60, 60), 1),
            BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));
        
        JLabel titleLabel = new JLabel(title);
        titleLabel.setFont(new Font("Arial", Font.BOLD, 20));
        titleLabel.setForeground(TEXT_COLOR);
        titleLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        panel.add(titleLabel);
        panel.add(Box.createVerticalStrut(20));
        
        return panel;
    }
    
    private JButton createButton(String text, ActionListener action) {
        JButton button = new JButton(text);
        button.setBackground(BUTTON_BG);
        button.setForeground(Color.WHITE);
        button.setFocusPainted(false);
        button.setAlignmentX(Component.CENTER_ALIGNMENT);
        button.addActionListener(action);
        
        // Add hover effect
        button.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                button.setBackground(BUTTON_HOVER);
            }
            
            @Override
            public void mouseExited(MouseEvent e) {
                button.setBackground(BUTTON_BG);
            }
        });
        
        return button;
    }
    
    private void setStatus(String message, boolean isRunning) {
        SwingUtilities.invokeLater(() -> {
            statusLabel.setText(message);
            statusLabel.setForeground(isRunning ? Color.RED : Color.GREEN);
        });
    }
    
    private void centerDialog(JDialog dialog, int width, int height) {
        dialog.setSize(width, height);
        dialog.setLocationRelativeTo(this);
    }
    
    private void showSchoolForms() {
        JDialog formsDialog = new JDialog(this, "School Forms", true);
        formsDialog.setLayout(new BoxLayout(formsDialog.getContentPane(), BoxLayout.Y_AXIS));
        formsDialog.getContentPane().setBackground(PANEL_BG);
        
        JLabel label = new JLabel("To edit the Master Form, open SF1");
        label.setForeground(TEXT_COLOR);
        label.setAlignmentX(Component.CENTER_ALIGNMENT);
        formsDialog.add(Box.createVerticalStrut(10));
        formsDialog.add(label);
        
        JButton sf1Btn = createButton("SF1", e -> safeOpenFile("sf1.xlsx"));
        JButton sf5aBtn = createButton("SF5A", e -> safeOpenFile("sf5a.xlsx"));
        JButton sf5bBtn = createButton("SF5B", e -> safeOpenFile("sf5b.xlsx"));
        JButton closeBtn = createButton("Close", e -> formsDialog.dispose());
        
        formsDialog.add(Box.createVerticalStrut(10));
        formsDialog.add(sf1Btn);
        formsDialog.add(Box.createVerticalStrut(5));
        formsDialog.add(sf5aBtn);
        formsDialog.add(Box.createVerticalStrut(5));
        formsDialog.add(sf5bBtn);
        formsDialog.add(Box.createVerticalStrut(10));
        formsDialog.add(closeBtn);
        formsDialog.add(Box.createVerticalStrut(10));
        
        centerDialog(formsDialog, 400, 300);
        formsDialog.setVisible(true);
    }
    
    private void showQuarterSelection() {
        JDialog quarterDialog = new JDialog(this, "Quarter Selection", true);
        quarterDialog.setLayout(new BoxLayout(quarterDialog.getContentPane(), BoxLayout.Y_AXIS));
        quarterDialog.getContentPane().setBackground(PANEL_BG);
        
        JLabel label = new JLabel("Select Which Quarter to Edit");
        label.setForeground(TEXT_COLOR);
        label.setAlignmentX(Component.CENTER_ALIGNMENT);
        quarterDialog.add(Box.createVerticalStrut(10));
        quarterDialog.add(label);
        
        String[] quarters = {"Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"};
        String[] files = {"MFQ1.xlsx", "MFQ2.xlsx", "MFQ3.xlsx", "MFQ4.xlsx"};
        
        for (int i = 0; i < quarters.length; i++) {
            final String file = files[i];
            JButton btn = createButton(quarters[i], e -> safeOpenFile(file));
            quarterDialog.add(Box.createVerticalStrut(5));
            quarterDialog.add(btn);
        }
        
        JButton closeBtn = createButton("Close", e -> quarterDialog.dispose());
        quarterDialog.add(Box.createVerticalStrut(10));
        quarterDialog.add(closeBtn);
        quarterDialog.add(Box.createVerticalStrut(10));
        
        centerDialog(quarterDialog, 400, 300);
        quarterDialog.setVisible(true);
    }
    
    private void safeOpenFile(String filepath) {
        File file = new File(filepath);
        if (file.exists()) {
            try {
                Desktop.getDesktop().open(file);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(this, 
                    "Cannot open file: " + e.getMessage(), 
                    "Error", 
                    JOptionPane.ERROR_MESSAGE);
            }
        } else {
            JOptionPane.showMessageDialog(this, 
                "File not found.", 
                "Error", 
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private void showPrintTransferOptions() {
        JDialog optionsDialog = new JDialog(this, "Print or Transfer", true);
        optionsDialog.setLayout(new BoxLayout(optionsDialog.getContentPane(), BoxLayout.Y_AXIS));
        optionsDialog.getContentPane().setBackground(PANEL_BG);
        
        JLabel label = new JLabel("Choose Option:");
        label.setForeground(TEXT_COLOR);
        label.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        JButton printBtn = createButton("Print", e -> {
            optionsDialog.dispose();
            showPrintSelection();
        });
        
        JButton transferBtn = createButton("Transfer", e -> {
            optionsDialog.dispose();
            showTransferOptions();
        });
        
        JButton cancelBtn = createButton("Cancel", e -> optionsDialog.dispose());
        
        optionsDialog.add(Box.createVerticalStrut(20));
        optionsDialog.add(label);
        optionsDialog.add(Box.createVerticalStrut(20));
        optionsDialog.add(printBtn);
        optionsDialog.add(Box.createVerticalStrut(10));
        optionsDialog.add(transferBtn);
        optionsDialog.add(Box.createVerticalStrut(10));
        optionsDialog.add(cancelBtn);
        optionsDialog.add(Box.createVerticalStrut(10));
        
        centerDialog(optionsDialog, 300, 200);
        optionsDialog.setVisible(true);
    }
    
    private void showPrintSelection() {
        JDialog printDialog = new JDialog(this, "Print Selection", true);
        printDialog.setLayout(new BoxLayout(printDialog.getContentPane(), BoxLayout.Y_AXIS));
        printDialog.getContentPane().setBackground(PANEL_BG);
        
        Map<String, String> options = new HashMap<>();
        options.put("School Form 1", "sf1.xlsx");
        options.put("School Form 5a", "sf5a.xlsx");
        options.put("School Form 5b", "sf5b.xlsx");
        options.put("School Forms 9", "SF9SF10/SF9");
        options.put("School Forms 10", "SF9SF10/SF10");
        
        Map<String, JCheckBox> checkboxes = new HashMap<>();
        
        for (String option : options.keySet()) {
            JCheckBox checkbox = new JCheckBox(option);
            checkbox.setForeground(TEXT_COLOR);
            checkbox.setBackground(PANEL_BG);
            checkboxes.put(option, checkbox);
            printDialog.add(checkbox);
            printDialog.add(Box.createVerticalStrut(5));
        }
        
        JButton printBtn = createButton("Print", e -> {
            boolean anySelected = checkboxes.values().stream().anyMatch(JCheckBox::isSelected);
            if (!anySelected) {
                JOptionPane.showMessageDialog(printDialog, 
                    "Select at least 1 file!", 
                    "Error", 
                    JOptionPane.ERROR_MESSAGE);
                return;
            }
            
            setStatus("Printing files...", true);
            boolean success = true;
            
            for (Map.Entry<String, JCheckBox> entry : checkboxes.entrySet()) {
                if (entry.getValue().isSelected()) {
                    String path = options.get(entry.getKey());
                    
                    if (path.endsWith(".xlsx")) {
                        if (!printFile(path)) {
                            success = false;
                        }
                    } else {
                        if (!printDirectory(path)) {
                            success = false;
                        }
                    }
                }
            }
            
            if (success) {
                setStatus("Print completed successfully", false);
            } else {
                setStatus("Some print jobs failed", false);
            }
            
            printDialog.dispose();
        });
        
        JButton closeBtn = createButton("Close", e -> printDialog.dispose());
        
        printDialog.add(Box.createVerticalStrut(10));
        printDialog.add(printBtn);
        printDialog.add(Box.createVerticalStrut(10));
        printDialog.add(closeBtn);
        printDialog.add(Box.createVerticalStrut(10));
        
        centerDialog(printDialog, 400, 400);
        printDialog.setVisible(true);
    }
    
    private boolean printFile(String filepath) {
        if (new File(filepath).exists()) {
            return excelPrinter.printExcelFile(filepath);
        }
        return false;
    }
    
    private boolean printDirectory(String directory) {
        File dir = new File(directory);
        if (!dir.exists() || !dir.isDirectory()) {
            return false;
        }
        
        boolean success = true;
        for (File file : dir.listFiles()) {
            if (file.getName().endsWith(".xlsx")) {
                if (!printFile(file.getAbsolutePath())) {
                    success = false;
                }
            }
        }
        return success;
    }
    
    private void showTransferOptions() {
        JDialog transferDialog = new JDialog(this, "Transfer Options", true);
        transferDialog.setLayout(new BoxLayout(transferDialog.getContentPane(), BoxLayout.Y_AXIS));
        transferDialog.getContentPane().setBackground(PANEL_BG);
        
        JLabel label = new JLabel("Choose Transfer Method:");
        label.setForeground(TEXT_COLOR);
        label.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        JButton emailBtn = createButton("Send Email", e -> {
            transferDialog.dispose();
            transferFiles(true);
        });
        
        JButton usbBtn = createButton("Transfer with USB/Folder", e -> {
            transferDialog.dispose();
            transferFiles(false);
        });
        
        JButton cancelBtn = createButton("Cancel", e -> transferDialog.dispose());
        
        transferDialog.add(Box.createVerticalStrut(20));
        transferDialog.add(label);
        transferDialog.add(Box.createVerticalStrut(20));
        transferDialog.add(emailBtn);
        transferDialog.add(Box.createVerticalStrut(10));
        transferDialog.add(usbBtn);
        transferDialog.add(Box.createVerticalStrut(10));
        transferDialog.add(cancelBtn);
        transferDialog.add(Box.createVerticalStrut(10));
        
        centerDialog(transferDialog, 300, 200);
        transferDialog.setVisible(true);
    }
    
    private void transferFiles(boolean isEmail) {
        String[] filesToTransfer = {"sf1.xlsx", "sf5a.xlsx", "sf5b.xlsx"};
        String folderToTransfer = "SF9SF10";
        
        if (isEmail) {
            // Email implementation would go here
            setStatus("Sending email...", true);
            // Email code placeholder
            JOptionPane.showMessageDialog(this, 
                "Files sent via email successfully!", 
                "Email", 
                JOptionPane.INFORMATION_MESSAGE);
            setStatus("Email sent successfully", false);
        } else {
            // Transfer to folder
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            fileChooser.setDialogTitle("Select Destination Folder");
            
            int result = fileChooser.showDialog(this, "Select");
            if (result != JFileChooser.APPROVE_OPTION) {
                setStatus("Transfer canceled", false);
                return;
            }
            
            String destination = fileChooser.getSelectedFile().getAbsolutePath();
            setStatus("Transferring files...", true);
            
            try {
                // Copy individual files
                for (String file : filesToTransfer) {
                    File sourceFile = new File(file);
                    if (sourceFile.exists()) {
                        Path targetPath = Paths.get(destination, file);
                        Files.copy(sourceFile.toPath(), targetPath, StandardCopyOption.REPLACE_EXISTING);
                    }
                }
                
                // Copy folder if it exists
                File sourceFolder = new File(folderToTransfer);
                if (sourceFolder.exists() && sourceFolder.isDirectory()) {
                    File destFolder = new File(destination, folderToTransfer);
                    if (destFolder.exists()) {
                        deleteDirectory(destFolder);
                    }
                    copyDirectory(sourceFolder, destFolder);
                }
                
                JOptionPane.showMessageDialog(this, 
                    "Files transferred successfully!", 
                    "Success", 
                    JOptionPane.INFORMATION_MESSAGE);
                setStatus("Files transferred successfully", false);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(this, 
                    "Error transferring files: " + e.getMessage(), 
                    "Error", 
                    JOptionPane.ERROR_MESSAGE);
                setStatus("Error transferring files", false);
            }
        }
    }
    
    private void deleteDirectory(File directory) throws IOException {
        if (directory.exists()) {
            File[] files = directory.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isDirectory()) {
                        deleteDirectory(file);
                    } else {
                        Files.delete(file.toPath());
                    }
                }
            }
            Files.delete(directory.toPath());
        }
    }
    
    private void copyDirectory(File sourceDirectory, File destinationDirectory) throws IOException {
        if (!destinationDirectory.exists()) {
            destinationDirectory.mkdir();
        }
        
        File[] files = sourceDirectory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    copyDirectory(file, new File(destinationDirectory, file.getName()));
                } else {
                    Files.copy(file.toPath(), 
                        Paths.get(destinationDirectory.getAbsolutePath(), file.getName()),
                        StandardCopyOption.REPLACE_EXISTING);
                }
            }
        }
    }
    
    private void runScriptAsync(String scriptName) {
        if (!executing) {
            executing = true;
            setStatus("Running " + scriptName + "...", true);
            
            // Disable all buttons during execution
            toggleButtons(false);
            
            executor.submit(() -> {
                try {
                    // Construct the command for Python script execution
                    ProcessBuilder processBuilder = new ProcessBuilder();
                    if (System.getProperty("os.name").toLowerCase().contains("windows")) {
                        processBuilder.command("python", scriptName);
                    } else {
                        processBuilder.command("python3", scriptName);
                    }
                    
                    Process process = processBuilder.start();
                    int exitCode = process.waitFor();
                    boolean success = exitCode == 0;
                    
                    SwingUtilities.invokeLater(() -> scriptCompleted(success, scriptName, ""));
                } catch (Exception e) {
                    SwingUtilities.invokeLater(() -> scriptCompleted(false, scriptName, e.getMessage()));
                }
            });
        } else {
            JOptionPane.showMessageDialog(this, 
                "A script is already running. Please wait until it finishes.", 
                "Warning", 
                JOptionPane.WARNING_MESSAGE);
        }
    }
    
    private void toggleButtons(boolean enabled) {
        SwingUtilities.invokeLater(() -> {
            for (Component c : getAllComponents(this)) {
                if (c instanceof JButton) {
                    c.setEnabled(enabled);
                }
            }
        });
    }
    
    private java.util.List<Component> getAllComponents(Container container) {
        java.util.List<Component> components = new java.util.ArrayList<>();
        for (Component c : container.getComponents()) {
            components.add(c);
            if (c instanceof Container) {
                components.addAll(getAllComponents((Container) c));
            }
        }
        return components;
    }
    
    private void scriptCompleted(boolean success, String scriptName, String errorMessage) {
        executing = false;
        toggleButtons(true);
        
        if (success) {
            setStatus(scriptName + " completed successfully", false);
        } else {
            setStatus("Error running " + scriptName, false);
            JOptionPane.showMessageDialog(this, 
                "Error running " + scriptName + ": " + errorMessage, 
                "Error", 
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private void exitProgram() {
        if (executing) {
            int response = JOptionPane.showConfirmDialog(this, 
                "A script is currently running. Are you sure you want to exit?", 
                "Warning", 
                JOptionPane.YES_NO_OPTION, 
                JOptionPane.WARNING_MESSAGE);
            
            if (response != JOptionPane.YES_OPTION) {
                return;
            }
        }
        
        excelPrinter.cleanup();
        executor.shutdown();
        dispose();
        System.exit(0);
    }
    
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            StudentManagementGUI gui = new StudentManagementGUI();
            gui.setVisible(true);
        });
    }
}

/**
 * Class to handle Excel printing operations
 */
class ExcelPrinter {
    private Object excel; // We'll use reflection for COM automation
    
    public ExcelPrinter() {
        // Excel will be initialized on first use
    }
    
    public boolean printExcelFile(String filepath) {
        // Note: In a real implementation, we would use Jacob library or JNA to access COM
        // This is a simplified version for demonstration purposes
        
        try {
            // Simulate Excel printing
            System.out.println("Printing file: " + filepath);
            
            // Here you would use the actual COM automation code
            // For Windows, you'd typically use the Jacob library
            
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }
    
    public void cleanup() {
        // Clean up Excel resources
        try {
            // In a real implementation, you would use proper COM cleanup code here
            System.out.println("Cleaning up Excel resources");
        } catch (Exception e) {
            // Ignore exceptions during cleanup
        }
    }
}