package main.java.SearchDocX;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

//https://stackoverflow.com/questions/50767762/unable-to-get-hyperlink-label-for-docx-file-using-apache-poi
//Look at to find link

//https://stackoverflow.com/questions/40912421/how-to-find-and-replace-text-in-word-files-both-doc-and-docx
//find and replace

public class FindFile {
	
	//private static final String OUTPUT = 
			//"C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files\\";
	
	public static void findFile(String path) {
		
		File directory = new File(path);
		File[] fileList = directory.listFiles();
		
		for (int i = 0; i < fileList.length; i++) {
			if (fileList[i].isFile()) {
				System.out.println("File " + fileList[i].getName());
			}
			else if (fileList[i].isDirectory()) {
				System.out.println("Directory " + fileList[i].getName());
			}
		} 
	}
	
	public static int getFileTotal(String directoryInput) {
		File directory = new File(directoryInput);
		File[] fileList = directory.listFiles();
		
		int fileCount = 0;
		
		for (int i = 0; i < fileList.length; i++) {
			if (fileList[i].isFile()) {
				fileCount++;
			}
		} 
		
		return fileCount;
	}
	
	public static void find_Remove_DocX(String path, String replacement) throws IOException, 
	InvalidFormatException {
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
			
			//Find and replace in tables->row->cell->paragraph
			List<XWPFTable> tableList = document.getTables();
			
			for (XWPFTable table : tableList) {
				List<XWPFTableRow> rowList = table.getRows();
				
				for (XWPFTableRow row: rowList) {
					List<XWPFTableCell> cellList = row.getTableCells();
					
					for(XWPFTableCell cell : cellList) {
						paragraphList = cell.getParagraphs();
						
						for(XWPFParagraph paragraph: paragraphList) {
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
					}
				}
			}
			
			//create new file
			//String newOutput = OUTPUT + "new" + file.getName();
			//System.out.println(newOutput);
			//document.write(new FileOutputStream(newOutput));
			
			//replace file
			document.write(new FileOutputStream(path));
			
			//close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
			} 
		
		catch (Exception ex) {
				ex.printStackTrace();
		}
	}
	
}//end of FindFile
