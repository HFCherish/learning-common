package learning.common.excel.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author hf_cherish
 * @date 2018/7/25
 */
public class FileUtils {
    public static void copyFileUsingStream(File source, File dest) throws IOException {
        FileInputStream inputStream = null;
        FileOutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(source);
            outputStream = new FileOutputStream(dest);

            byte[] buffer = new byte[1024];

            int length;

            while ((length = inputStream.read(buffer)) > 0) {
                outputStream.write(buffer, 0, length);
            }
        } finally {
            inputStream.close();
            outputStream.close();
        }
    }

    public static File getFile(String name) {
        String file = FileUtils.class.getResource(name).getFile();
        return new File(file);
    }
}
