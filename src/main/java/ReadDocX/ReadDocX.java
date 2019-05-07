/**
 * Read FIG document
 * Reference: https://www.devglan.com/corejava/parsing-word-document-example
 */

package main.java.ReadDocX;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class ReadDocX {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println("Class ReadDocx.");
		readDocXFile();
	}
	
	public static void readDocXFile() {
		String path = "C:/Users/CatherineKhongsaly/Documents/Test/Test Template.docx";
		
		try {
			
			File file = new File(path);
			
			FileInputStream fis = new FileInputStream(file);
			
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));

			XWPFWordExtractor extractor = new XWPFWordExtractor(document);
			
			System.out.println(extractor.getText());
			
			//close FileInputSTream and XWPFDocument
			fis.close();
			document.close();
			extractor.close();
			
		} catch (Exception e) {
			
			e.printStackTrace();
			
		}
	}

}
