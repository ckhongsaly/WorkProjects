/**
 * Read FIG document
 * Reference: https://www.devglan.com/corejava/parsing-word-document-example
 */

package main.java.ReadDocX;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class ReadDocX {

	public static void readDocXFile(String path) {
		
		try {
			
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
			XWPFWordExtractor extractor = new XWPFWordExtractor(document);
			
			System.out.println(extractor.getText());
			
			//close FileInputSTream, XWPFDocument, and WordExtractor
			fis.close();
			document.close();
			extractor.close();
			
		} catch (Exception e) {
			
			e.printStackTrace();
			
		}
	}// end of readDocXFile
	
	public static void readHeaderFooterDocX (String path) {
		
		try {
			
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
			XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document);
			
			//Read header
			XWPFHeader header = policy.getDefaultHeader();
			
			if (header != null) {
				System.out.println(header.getText());
			}
			
			//Read footer
			XWPFFooter footer = policy.getDefaultFooter();
			
			if (footer != null) {
				System.out.println(footer.getText());
			}
			
			//close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
			
		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}//end of readHeaderFooterDocX
	
	public static void readParagraphDocX(String path) {
		try {
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));

			List<XWPFParagraph> paragraphList = document.getParagraphs();

			for (XWPFParagraph paragraph : paragraphList) {

				System.out.println(paragraph.getText());
				System.out.println(paragraph.getAlignment());
				System.out.print(paragraph.getRuns().size());
				System.out.println(paragraph.getStyle());

				// Returns numbering format for this paragraph, eg bullet or lowerLetter.
				System.out.println(paragraph.getNumFmt());
				System.out.println(paragraph.getAlignment());

				System.out.println(paragraph.isWordWrapped());

				System.out.println("********************************************************************");
				
				//close FileInputSTream, XWPFDocument
				fis.close();
				document.close();
				
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	
	public static void readTableDocX(String path) {
		try {
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
			Iterator<IBodyElement> bodyElementIterator = document.getBodyElementsIterator();
			while (bodyElementIterator.hasNext()) {
				IBodyElement element = bodyElementIterator.next();

				if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
					List<XWPFTable> tableList = element.getBody().getTables();
					for (XWPFTable table : tableList) {
						System.out.println("Total Number of Rows of Table:" + table.getNumberOfRows());
						for (int i = 0; i < table.getRows().size(); i++) {

							for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
								System.out.println(table.getRow(i).getCell(j).getText());
							}
						}
					}
				}
			}
			
			//close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
			
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	
	public static void readStyle(String path) {
		try {
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));

			List<XWPFParagraph> paragraphList = document.getParagraphs();

			for (XWPFParagraph paragraph : paragraphList) {

				for (XWPFRun rn : paragraph.getRuns()) {

					System.out.println(rn.isBold());
					System.out.println(rn.isHighlighted());
					System.out.println(rn.isCapitalized());
					System.out.println(rn.getFontSize());
				}

				System.out.println("********************************************************************");
			}
			
			//close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
			
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	
	public static void readImage(String path) {
		try {
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
			List<XWPFPictureData> pic = document.getAllPictures();
			if (!pic.isEmpty()) {
				System.out.print(pic.get(0).getPictureType());
				System.out.print(pic.get(0).getData());
			}
			
			//close FileInputSTream, XWPFDocument
			fis.close();
			document.close();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	
}//end of class

