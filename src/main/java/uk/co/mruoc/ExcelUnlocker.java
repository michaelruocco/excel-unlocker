package uk.co.mruoc;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
//import java.util.List;

//import static java.io.File.separator;

public class ExcelUnlocker {

    //private final FileGetter fileGetter = new FileGetter();
    //private final DirectoryCreator directoryCreator = new DirectoryCreator();

    /*private final String inputDirectoryPath;
    private final String outputDirectoryPath;
    private final String password;*/

    /*private ExcelUnlocker(ExcelUnlockerBuilder builder) {
        this.inputDirectoryPath = builder.inputDirectoryPath;
        this.outputDirectoryPath = builder.outputDirectoryPath;
        this.password = builder.password;
    }

    public void unlockFiles() {
        List<File> files = fileGetter.getFilesInDirectory(inputDirectoryPath);
        File outputDirectory = directoryCreator.createDirectory(outputDirectoryPath);
        for (File file : files) {
            unlock(file, password, outputDirectory);
        }
    }*/

    public void unlock(File inputFile, String password, File outputFile) {
        if (inputFile.isFile()) {
            /*Workbook workbook = */openProtectedWorkbook(inputFile, password);
            //createUnlockedCopyOfWorkbook(workbook, outputFile);
        }
    }

    private void /*Workbook*/ openProtectedWorkbook(File file, String password) {
        try {
            /*POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);
            decryptor.verifyPassword(password);
            XSSFWorkbook workbook = new XSSFWorkbook(decryptor.getDataStream(fs));*/

            LoadOptions loadOptions = new LoadOptions(FileFormatType.XLSX);
            //loadOptions.setPassword(password);
            com.aspose.cells.Workbook workbook1 = new com.aspose.cells.Workbook(new FileInputStream(file), loadOptions);
            workbook1.unprotect(password);
            workbook1.save(file.getAbsolutePath());
            //return workbook;

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void createUnlockedCopyOfWorkbook(Workbook workbook, File outputFile) {
        try {
            String unlockedPath = outputFile.getAbsolutePath();
            FileOutputStream stream = new FileOutputStream(unlockedPath);
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /*public static class ExcelUnlockerBuilder {

        private String inputDirectoryPath;
        private String outputDirectoryPath;
        private String password;

        public ExcelUnlockerBuilder setInputDirectoryPath(String inputDirectoryPath) {
            this.inputDirectoryPath = inputDirectoryPath;
            return this;
        }

        public ExcelUnlockerBuilder setOutputDirectoryPath(String outputDirectoryPath) {
            this.outputDirectoryPath = outputDirectoryPath;
            return this;
        }

        public ExcelUnlockerBuilder setPassword(String password) {
            this.password = password;
            return this;
        }

        public ExcelUnlocker build() {
            return new ExcelUnlocker(this);
        }

    }*/

}
