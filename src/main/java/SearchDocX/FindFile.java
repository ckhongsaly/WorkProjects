package main.java.SearchDocX;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class FindFile {
	
	public static void findFile(File directory) {
		File[] list = directory.listFiles();
		
		if(list != null) {
			for (File file : list) {
				if (file.isDirectory()) {
					System.out.println(file.getParentFile());
				}
			}
		}
	}
	
	public static void find_Replace_DocX(String path, String replacement) throws IOException, InvalidFormatException {
		try {
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
			
			//Find and replace in paragraph
			
			List<XWPFParagraph> paragraphList = document.getParagraphs();
			
			for (XWPFParagraph paragraph : paragraphList) {
				List<XWPFRun> runs = paragraph.getRuns();
				if(runs != null) {
					for (XWPFRun r: runs) {
						String text = r.getText(0);
						if (text != null && text.contains(replacement)) {
							text = text.replace(replacement, ""); //remove text
							r.setText(text,0);
						}
					}
				}
			}
			
			
			
			
			//close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
			} 
		catch (Exception ex) {
				ex.printStackTrace();
			}
	}
	
}//end of FindFile
