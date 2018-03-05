package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

public class PropertiesUtils {

    public static Map<String, String> readPropertiesToExcel(File file) throws Exception {
    	Map<String, String> retObj = new LinkedHashMap<String, String>();
        Properties props = new Properties();
        InputStream stream = new FileInputStream(file);
        props.load(stream);
		for (String key :  props.stringPropertyNames()) {
			if(!key.startsWith("word")) {
				retObj.put(key, props.getProperty(key));
			}
		}
        return retObj;
    }
    
    public static Map<String, String> readPropertiesToWord(File file) throws Exception {
    	Map<String, String> retObj = new LinkedHashMap<String, String>();
        Properties props = new Properties();
        InputStream stream = new FileInputStream(file);
        props.load(stream);
		for (String key :  props.stringPropertyNames()) {
			if(key.startsWith("word")) {
				retObj.put(key, props.getProperty(key));
			}
		}
        return retObj;
    }
}
