package workExcel;

import java.awt.Desktop;
import java.awt.EventQueue;
import java.awt.Color;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.JScrollPane;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableRowSorter;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelProject {

    private JFrame frame;
    private JTable table;
    private DefaultTableModel tableModel;
    private TableRowSorter<DefaultTableModel> sorter;

    public static void main(String[] args) {
        EventQueue.invokeLater(() -> {
            try {
                ExcelProject window = new ExcelProject();
                window.frame.setVisible(true);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

    public ExcelProject() {
        initialize();
    }

    private void initialize() {
        frame = new JFrame();
        frame.setBounds(100, 100, 1058, 513);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().setLayout(null);
        frame.setTitle("Excel Analysis");
        
        
      

        tableModel = new DefaultTableModel();
        tableModel.setColumnIdentifiers(new String[]{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"});

        table = new JTable(tableModel);
        table.getTableHeader().setBackground(Color.CYAN);

        JScrollPane scrollPane = new JScrollPane(table);
        scrollPane.setBounds(10, 11, 1000, 403);
        frame.getContentPane().add(scrollPane);

        sorter = new TableRowSorter<>(tableModel);
        table.setRowSorter(sorter);

        JButton btnNewButton = new JButton("Import");
        btnNewButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                try {
                    importExcel();
                } catch (FileNotFoundException e1) {
                    e1.printStackTrace();
                }
            }
        });
        btnNewButton.setBounds(10, 425, 89, 23);
        frame.getContentPane().add(btnNewButton);

        JButton btnHtml = new JButton("HTML File");
        btnHtml.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                int result = JOptionPane.showConfirmDialog(
                        null,
                        "Do you want to open the HTML file?",
                        "Open HTML File",
                        JOptionPane.YES_NO_OPTION
                );

                if (result == JOptionPane.YES_OPTION) {
                    createHtmlFiles();
                }

                JOptionPane.showMessageDialog(null, "HTML files created successfully.");
            }

            private void createHtmlFiles() {
                try {
                    File combinedHtmlFile = new File("combined.html");
                    BufferedWriter writer = new BufferedWriter(new FileWriter(combinedHtmlFile));

                    // Write HTML header
                    writer.write("<html><head><title>Combined Records</title></head><body>");
                    // Write table header
                    writer.write("<table border=\"1\"><tr>");
                    for (int columnIndex = 0; columnIndex < tableModel.getColumnCount(); columnIndex++) {
                        writer.write("<th>" + tableModel.getColumnName(columnIndex) + "</th>");
                    }
                    writer.write("</tr>");

                    // Write rows for all columns
                    for (int rowIndex = 0; rowIndex < tableModel.getRowCount(); rowIndex++) {
                        writer.write("<tr>");
                        for (int columnIndex = 0; columnIndex < tableModel.getColumnCount(); columnIndex++) {
                            writer.write("<td>" + tableModel.getValueAt(rowIndex, columnIndex) + "</td>");
                        }
                        writer.write("</tr>");
                    }

                    // Write HTML footer
                    writer.write("</table></body></html>");
                    writer.close();

                    // Open the combined HTML file in the default web browser
                    openHtmlFileInBrowser(combinedHtmlFile.getAbsolutePath());
                } catch (IOException e) {
                    JOptionPane.showMessageDialog(null, "Error creating combined HTML file: " + e.getMessage());
                }
            }

            private void openHtmlFileInBrowser(String htmlFilePath) {
                try {
                    File file = new File(htmlFilePath);
                    if (Desktop.isDesktopSupported()) {
                        Desktop.getDesktop().browse(file.toURI());
                    } else {
                        JOptionPane.showMessageDialog(null, "Desktop not supported. Please open the HTML file manually: " + htmlFilePath);
                    }
                } catch (IOException e) {
                    JOptionPane.showMessageDialog(null, "Error opening HTML file: " + e.getMessage());
                }
            }
        });
        btnHtml.setBounds(110, 425, 220, 23);
        frame.getContentPane().add(btnHtml);
    }

    private void importExcel() throws FileNotFoundException {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Please select the Excel file");
        int res = fileChooser.showOpenDialog(null);

        if (res == JFileChooser.APPROVE_OPTION) {
            String excelPath = fileChooser.getSelectedFile().getAbsolutePath();
            File file = new File(excelPath);
            FileInputStream inputStream = new FileInputStream(file);

            try (Workbook workbook = new XSSFWorkbook(inputStream)) {
                Sheet sheet = workbook.getSheetAt(0);
                List<Object[]> excelData = new ArrayList<>();
                    
                

                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        List<Object> rowData = new ArrayList<>();
                        for (int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++) {
                            Cell cell = row.getCell(columnIndex);
                            if (cell != null) {
                                rowData.add(cell.toString());
                            } else {
                                rowData.add("");
                            }
                        }
                        excelData.add(rowData.toArray());
                        
                    }
                }
                updateTableModel(excelData);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, e.getMessage());
            }
            
        }
    
    }

    private void updateTableModel(List<Object[]> data) {
        tableModel.setRowCount(0);
        for (Object[] rowData : data) {
            tableModel.addRow(rowData);
        }

    }

}


