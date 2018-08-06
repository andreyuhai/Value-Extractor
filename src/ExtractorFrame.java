
import java.awt.Desktop;
import java.awt.Image;
import java.awt.Toolkit;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.net.URL;
import java.nio.file.FileAlreadyExistsException;
import java.util.ArrayList;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author burak
 */
public class ExtractorFrame extends javax.swing.JFrame {

    /**
     * Creates new form ExtractorFrame
     */
    private File selectedFiles[];
    public ExtractorFrame() {
        initComponents();
        this.setVisible(true);
        this.setLocationRelativeTo(null);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        selectedFiles_Label = new javax.swing.JLabel();
        selectedFiles_TF = new javax.swing.JTextField();
        browse_Button = new javax.swing.JButton();
        extract_Button = new javax.swing.JButton();
        menuBar = new javax.swing.JMenuBar();
        file_Menu = new javax.swing.JMenu();
        exit_MenuItem = new javax.swing.JMenuItem();
        help_Menu = new javax.swing.JMenu();
        howToUse_MenuItem = new javax.swing.JMenuItem();
        about_MenuItem = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Value Extractor v0.0.1");
        setIconImage(getIconImage());
        setResizable(false);

        selectedFiles_Label.setText("Files Selected :");

        browse_Button.setText("Browse");
        browse_Button.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                browse_ButtonActionPerformed(evt);
            }
        });

        extract_Button.setText("Extract");
        extract_Button.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                extract_ButtonActionPerformed(evt);
            }
        });

        file_Menu.setText("File");
        file_Menu.setToolTipText("");

        exit_MenuItem.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_F4, java.awt.event.InputEvent.ALT_MASK));
        exit_MenuItem.setText("Exit");
        exit_MenuItem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exit_MenuItemActionPerformed(evt);
            }
        });
        file_Menu.add(exit_MenuItem);

        menuBar.add(file_Menu);

        help_Menu.setText("Help");

        howToUse_MenuItem.setText("How to use?");
        howToUse_MenuItem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                howToUse_MenuItemActionPerformed(evt);
            }
        });
        help_Menu.add(howToUse_MenuItem);

        about_MenuItem.setText("About");
        about_MenuItem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                about_MenuItemActionPerformed(evt);
            }
        });
        help_Menu.add(about_MenuItem);

        menuBar.add(help_Menu);

        setJMenuBar(menuBar);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
                    .addComponent(extract_Button)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(selectedFiles_Label)
                        .addGap(12, 12, 12)
                        .addComponent(selectedFiles_TF, javax.swing.GroupLayout.PREFERRED_SIZE, 185, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(browse_Button)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
                    .addComponent(browse_Button, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(selectedFiles_TF, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(selectedFiles_Label))
                .addGap(18, 18, 18)
                .addComponent(extract_Button)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void browse_ButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_browse_ButtonActionPerformed
        // TODO add your handling code here:
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setAcceptAllFileFilterUsed(false);
        
        fileChooser.setFileFilter(new FileNameExtensionFilter("Text Dosyasi (.txt)", "txt"));
        
        int result = fileChooser.showOpenDialog(this);
        if(result == JFileChooser.APPROVE_OPTION) {
            String selectedFileNames = "";
            selectedFiles = fileChooser.getSelectedFiles();
            for(int i = 0; i < selectedFiles.length; i++) {
                selectedFileNames += selectedFiles[i].getName() + " ";
            }
            selectedFiles_TF.setText(selectedFileNames);
        }
    }//GEN-LAST:event_browse_ButtonActionPerformed

    private void howToUse_MenuItemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_howToUse_MenuItemActionPerformed
        // TODO add your handling code here:
        JOptionPane.showMessageDialog(this, "1. Browse and select the file(s) from which the values to be extracted.\n"
                + "2. Then extract the values.\n"
                + "\nAn excel file with the values will be created for each selected text file.\n"
                + "Note: This application works only on specifically formatted text files.","How to use?",JOptionPane.INFORMATION_MESSAGE);
    }//GEN-LAST:event_howToUse_MenuItemActionPerformed

    private void about_MenuItemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_about_MenuItemActionPerformed
        // TODO add your handling code here:
        openWebpage("http://blog.andreyuhai.com");
    }//GEN-LAST:event_about_MenuItemActionPerformed

    private void exit_MenuItemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exit_MenuItemActionPerformed
        // TODO add your handling code here:
        this.dispose();
    }//GEN-LAST:event_exit_MenuItemActionPerformed

    private void extract_ButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_extract_ButtonActionPerformed
        // TODO add your handling code here:
        try{
            for(int i = 0; i < selectedFiles.length; i++) {
                ArrayList valuesList = readFileAsList(selectedFiles[i]);
                write_toExcel(valuesList, selectedFiles[i]);
            }
            JOptionPane.showMessageDialog(this,"Successfully extracted values from "
                    +Integer.toString(selectedFiles.length) + " file(s).", "Extraction Succesful",JOptionPane.INFORMATION_MESSAGE);
            selectedFiles = null;
            selectedFiles_TF.setText("");
        }catch(NullPointerException e) {
            JOptionPane.showMessageDialog(this,"You didn't choose any file!", "Warning!",JOptionPane.WARNING_MESSAGE);
        }
    }//GEN-LAST:event_extract_ButtonActionPerformed

    public static void openWebpage(String urlString) {
        try {
            Desktop.getDesktop().browse(new URL(urlString).toURI());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    private ArrayList readFileAsList(File file) {
        boolean containsValue = false;
        BufferedReader reader = null;
        ArrayList<String> valuesList = new ArrayList<String>();
        
        try {
            reader = new BufferedReader(new FileReader(file));
            String line = null;
            String values[];
            
            
            while ((line = reader.readLine()) != null) {
               
                if(line.toLowerCase().contains("value")) {
                    containsValue = true;
                    continue;
                }
                
                
                if(containsValue) {
                    if(line.isEmpty()) {
                        containsValue = false;
                        break;
                    }
                    valuesList.add(line.split(";")[3].trim());
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (reader != null) {
                    reader.close();
                }
            } catch (IOException e) {
            }
        }
        return valuesList;
    }
    
    private void write_toExcel(ArrayList arrayList, File fileName) {
        
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("First Sheet");
        for(int i = 0; i < arrayList.size(); i++) {
            HSSFRow row = sheet.createRow(i);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(arrayList.get(i).toString());
        }
        
        try{
            workbook.write(new FileOutputStream(fileName.getCanonicalPath().replace(".txt", "-extracted.xls")));
            workbook.close();
        }catch(FileAlreadyExistsException e) {
            JOptionPane.showMessageDialog(this,"file already exists", "file",JOptionPane.INFORMATION_MESSAGE);
        }catch(IOException e){
            e.printStackTrace();
        }
        
    }
    
    public Image getIconImage() {
//        ImageIcon icon = new ImageIcon("images/icon3.png");
//        System.out.println(getClass().getClassLoader().getResource(".").toString());
        return Toolkit.getDefaultToolkit().getImage(getClass().getResource("images/icon3.png"));
    }
    
    
    
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JMenuItem about_MenuItem;
    private javax.swing.JButton browse_Button;
    private javax.swing.JMenuItem exit_MenuItem;
    private javax.swing.JButton extract_Button;
    private javax.swing.JMenu file_Menu;
    private javax.swing.JMenu help_Menu;
    private javax.swing.JMenuItem howToUse_MenuItem;
    private javax.swing.JMenuBar menuBar;
    private javax.swing.JLabel selectedFiles_Label;
    private javax.swing.JTextField selectedFiles_TF;
    // End of variables declaration//GEN-END:variables
}