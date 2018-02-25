package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Properties;

public class PropertiesUtils {

    public static Properties readProperties(File file) throws Exception {

        Properties retObj = new Properties();
        InputStream stream = new FileInputStream(file);
        retObj.load(stream);
        return retObj;
    }
}
