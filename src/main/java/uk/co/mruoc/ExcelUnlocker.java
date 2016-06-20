package uk.co.mruoc;

import java.io.File;
import java.io.FileInputStream;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;

public class ExcelUnlocker {

    public void createUnprotectedCopy(File protectedFile, String password, String unprotectedCopyPath) {
        Workbook workbook = openProtectedWorkbook(protectedFile, password);
        createUnprotectedCopy(workbook, unprotectedCopyPath);
    }

    private Workbook openProtectedWorkbook(File file, String password) {
        try {
            LoadOptions loadOptions = new LoadOptions(FileFormatType.XLSX);
            Workbook workbook = new Workbook(new FileInputStream(file), loadOptions);
            workbook.unprotect(password);
            return workbook;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void createUnprotectedCopy(Workbook workbookToCopy, String pathToCopyTo) {
        try {
            Workbook target = new Workbook();
            target.copy(workbookToCopy);
            target.save(pathToCopyTo.toString());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

}
