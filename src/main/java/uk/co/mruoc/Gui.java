package uk.co.mruoc;

import static javax.swing.JOptionPane.*;
import static javax.swing.JOptionPane.showMessageDialog;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.*;

import org.apache.commons.io.FilenameUtils;


public class Gui extends JFrame {

    public Gui() {
        super("Excel Unlocker");

        JMenuBar menuBar = new JMenuBar();
        JMenu file = new JMenu("File");
        JMenuItem open = new JMenuItem("Open");

        file.add(open);
        menuBar.add(file);

        setJMenuBar(menuBar);

        OpenFileListener openFileListener = new OpenFileListener(this);
        open.addActionListener(openFileListener);

        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setPreferredSize(new Dimension(200, 200));
        setVisible(true);
        pack();
    }

    public static class OpenFileListener implements ActionListener {

        private final ExcelUnlocker excelUnlocker = new ExcelUnlocker();
        private JFrame frame;

        public OpenFileListener(JFrame frame) {
            this.frame = frame;
        }

        @Override
        public void actionPerformed(ActionEvent e) {
            File[] files = getSelectedFiles();
            String password = getPassword();
            int unprotectedCount = 0;
            for (File file : files)
                if (unprotect(file, password))
                    unprotectedCount++;
            showCompletionMessage(files, unprotectedCount);
        }

        private File[] getSelectedFiles() {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            chooser.showOpenDialog(frame);
            return chooser.getSelectedFiles();
        }

        private String getPassword() {
            return showInputDialog(frame, "Enter Password");
        }

        private boolean unprotect(File file, String password) {
            try {
                System.out.println("attempting to unprotect file " + file.getAbsolutePath());
                String outputFilePath = getOutputPath(file);
                excelUnlocker.createUnprotectedCopy(file, password, outputFilePath);
                System.out.println("unprotected copy created at " + outputFilePath);
                return true;
            } catch (Exception ex) {
                ex.printStackTrace();
                String message = "Error unlocking " + file.getAbsolutePath() + " : " + ex.getMessage();
                showMessageDialog(frame, message, "Error", ERROR_MESSAGE);
                return false;
            }
        }

        private String getOutputPath(File inputFile) {
            String path = inputFile.getParent();
            String name = FilenameUtils.getBaseName(inputFile.getName());
            String extension = FilenameUtils.getExtension(inputFile.getName());
            path = path + File.separator + name + "-unprotected." + extension;
            return path;
        }

        private void showCompletionMessage(File[] files, int unprotectedCount) {
            String message = "unprotected " + unprotectedCount + " of " + files.length + " files";
            showMessageDialog(frame, message);
        }

    }

}
