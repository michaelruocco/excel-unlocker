package uk.co.mruoc;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class FileGetter {

    public List<File> getFilesInDirectory(String path) {
        File directory = new File(path);
        File[] files = directory.listFiles();
        return new ArrayList<>(Arrays.asList(files));
    }

}
