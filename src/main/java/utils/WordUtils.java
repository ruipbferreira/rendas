package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import model.Fraction;

public class WordUtils {

	public static void generateWord(List<Fraction> fractions, File templatePath, String outputPath, Map<String, String> properties) throws Exception {
		String filePathOut = outputPath;
		POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(templatePath));
		for (Fraction fraction : fractions) {
			HWPFDocument doc = new HWPFDocument(fs);
			// wordnome
			doc = replaceText(doc, properties.get("wordnome"), fraction.getName());
			// wordnif
			doc = replaceText(doc, properties.get("wordnif"), fraction.getNif());
			// wordvalor
			doc = replaceText(doc, properties.get("wordvalor"), fraction.getValue());
			// wordmorada
			doc = replaceText(doc, properties.get("wordmorada"), fraction.getAddress());
			// wordfreguesia
			doc = replaceText(doc, properties.get("wordfreguesia"), fraction.getAddressCode());
			// wordartigo
			doc = replaceText(doc, properties.get("wordartigo"), fraction.getArticleCode());
			// wordfraccao
			doc = replaceText(doc, properties.get("wordfraccao"), fraction.getFractionCode());
			// wordquotaparte
			doc = replaceText(doc, properties.get("wordquotaparte"), fraction.getShare());
			String finalFilePathOut = filePathOut + System.getProperty("file.separator") + fraction.getName() + ".doc";
			saveWord(finalFilePathOut, doc);
		}
	}

	private static HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText){
		Range r1 = doc.getRange(); 

		for (int i = 0; i < r1.numSections(); ++i ) { 
			Section s = r1.getSection(i); 
			for (int x = 0; x < s.numParagraphs(); x++) { 
				Paragraph p = s.getParagraph(x); 
				for (int z = 0; z < p.numCharacterRuns(); z++) { 
					CharacterRun run = p.getCharacterRun(z); 
					String text = run.text();
					if(text.contains(findText)) {
						run.replaceText(findText, replaceText);
					} 
				}
			}
		} 
		return doc;
	}

	private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException{
		FileOutputStream out = null;
		try{
			out = new FileOutputStream(filePath);
			doc.write(out);
		}
		finally{
			out.close();
		}
	}
}
