package uk.co.mruoc;

import org.apache.commons.io.FilenameUtils;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

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

        private JFrame frame;

        public OpenFileListener(JFrame frame) {
            this.frame = frame;
        }

        @Override
        public void actionPerformed(ActionEvent e) {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            chooser.showOpenDialog(frame);

            File[] files = chooser.getSelectedFiles();

            String password = JOptionPane.showInputDialog(frame, "Enter Password");

            try {
                ExcelUnlocker excelUnlocker = new ExcelUnlocker();
                for (File file : files) {
                    System.out.println("file " + file.getName());
                    File outputFile = new File(getOutputPath(file));
                    System.out.println("output " + outputFile.getAbsolutePath());
                    excelUnlocker.unlock(file, password, outputFile);
                }

                JOptionPane.showMessageDialog(frame, files.length + " files unlocked successfully");
            } catch (NullPointerException ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(frame, "Error unlocking files", "Error", JOptionPane.ERROR_MESSAGE);
            }
        }

        private String getOutputPath(File inputFile) {
            String path = inputFile.getParent();
            String name = FilenameUtils.getBaseName(inputFile.getName());
            String extension = FilenameUtils.getExtension(inputFile.getName());
            path = path + File.separator + name + "-unlocked." + extension;
            return path;
        }

    }

}
