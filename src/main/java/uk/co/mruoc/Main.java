package uk.co.mruoc;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import javax.swing.*;

import java.io.File;
import java.io.IOException;

import static javax.swing.SwingUtilities.invokeLater;
import static uk.co.mruoc.ExcelUnlocker.*;

public class Main {

    public static void main(String[] args) {
        /*SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new Gui();
            }
        });*/
        /*ExcelUnlocker unlocker = new ExcelUnlockerBuilder()
                .setInputDirectoryPath("input")
                .setOutputDirectoryPath("output")
                .setPassword("test")
                .build();

        unlocker.unlockFiles();*/
        ExcelUnlocker unlocker = new ExcelUnlocker();
        unlocker.unlock(new File("input/Book1a.xlsx"), "test", new File("input/Book1a-unprotected-temp.xlsx"));
    }

}
