package uk.co.mruoc;

import java.io.File;

public class DirectoryCreator {

    public File createDirectory(String path) {
        File file = new File(path);

        if (!file.exists() && file.mkdirs()) {
            return file;
        }

        if (!file.isDirectory())
            throw new RuntimeException(path + " already exists and is not a directory");

        return file;
    }

}
