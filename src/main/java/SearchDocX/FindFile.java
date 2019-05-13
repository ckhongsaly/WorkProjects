package main.java.SearchDocX;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

//https://www.codota.com/code/java/classes/org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun

//https://github.com/Sayi/poi-tl/blob/master/src/main/java/com/deepoove/poi/XWPFParagraphWrapper.java#L49

//https://stackoverflow.com/questions/35088893/apache-poi-remove-cthyperlink-low-level-code/35101440#35101440

public class FindFile {
	
	private static XWPFHyperlink tempLink;
	
	/**
	 * List out files in a path
	 * @param path
	 */
	public static void find_File(String path) {
		
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
	}//findFile
	
	/**
	 * Return total number of files in path
	 * @param path
	 * @return fileCount
	 */
	public static int get_File_Total(String path) {
		File directory = new File(path);
		File[] fileList = directory.listFiles();
		
		int fileCount = 0;
		
		for (int i = 0; i < fileList.length; i++) {
			if (fileList[i].isFile()) {
				fileCount++;
			}
		} 
		
		return fileCount;
	}//end getFileTotal
	
	/**
	 * Description: Search path for file, search docx for replacement and then update replacement
	 * to newValue. Upload new file.
	 * @param path
	 * @param replacement
	 * @param newValue
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public static void find_Remove_DocX(String path, String replacement, String newValue) throws IOException, 
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
							text = text.replace(replacement, newValue); //remove text
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
	}//end find_Remove_DocX
	
	/**
	 * Description: Search path for file, search links in docx for replacement and 
	 * then update replacement to newValue. Upload new file.
	 * @param path
	 * @param replacement
	 * @param newValue
	 */
	public static void find_RemoveLink_DocX(String path, String replacement, String newValue) {
		
		boolean setup = true;
		
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
						
						if(r instanceof XWPFHyperlinkRun) {
							XWPFHyperlink link = ((XWPFHyperlinkRun) r).getHyperlink(document);
							
							if(setup) {
								XWPFHyperlink tempLink = link;
								setup = false;
							}
							if(link != null && link.getURL().contains(replacement) && tempLink != link) {
								//do something
							}
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
									
									if(r instanceof XWPFHyperlinkRun) {
										XWPFHyperlink link = ((XWPFHyperlinkRun) r).getHyperlink(document);
										
										if(setup) {
											XWPFHyperlink tempLink = link;
											setup = false;
										}
										if(link != null && link.getURL().contains(replacement) && tempLink != link) {
											//Create new link with replacement, get text, and create relationship
											String newUrl = link.getURL().replaceAll(replacement, newValue);
											String newText = r.text();
											String id = r.getDocument().getPackagePart()
													.addExternalRelationship(newUrl, XWPFRelation.HYPERLINK.getRelation()).getId();
											
											//Binds the link to the relationship
											CTHyperlink newHyperlink = paragraph.getCTP().addNewHyperlink();
											newHyperlink.set(((XWPFHyperlinkRun) r).getCTHyperlink());
											newHyperlink.setId(id);
									        
											//Creates the linked text
											CTText linkedText = CTText.Factory.newInstance();
											linkedText.setStringValue(newText);
									        
											//Creates a XML word processing wrapper for Run
											CTR ctr = r.getCTR();
											ctr.setTArray(new CTText[] {linkedText});
									        
									        //Style
											CTRPr rprC = ctr.addNewRPr();
											CTColor color = CTColor.Factory.newInstance();
											rprC.setColor(color);
											CTRPr rpr_u = ctr.addNewRPr();
											rpr_u.addNewU().setVal(STUnderline.SINGLE);

									        //remove hyperlink
					                        //paragraph.removeRun(0);
											int pos = r.getTextPosition();
											IRunBody parent = r.getParent();
					                        
					                        if (parent instanceof XWPFParagraph) {
					                            ((XWPFParagraph) parent).removeRun(((XWPFParagraph) parent).getRuns().indexOf(pos));
					                        } else {
					                            throw new IllegalStateException("this should not happend");
					                        }
					                        
					                        //add new hyperlink
									        r = new XWPFHyperlinkRun(newHyperlink, ctr, r.getParent());
										}
									}	
								}
							}
						}
					}
				}
			}
			
			//replace file
			document.write(new FileOutputStream(path));
			
			//close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
			} 
		
		catch (Exception ex) {
				ex.printStackTrace();
		}
	}//end of find_RemoveLink_DocX
	
	/**
	 * Description: Locate text in path and return value
	 * @param path
	 * @param replacement
	 * @return return total text found in docx
	 */
	public static int find_TextDocX(String path, String replacement) {
		int countText = 0;
		
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
							countText++;
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
										countText++;
									}
								}
							}
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
		
		return countText;
	}//end of find_TextDocX
	
	/**
	 * Description: Locate replacement in path and return value
	 * @param path
	 * @param replacement
	 * @return return total text found in link for docx
	 */
	public static int find_LinkDocX(String path, String replacement) {
		
		int linkCounter = 0;
		boolean setup = true;
		
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
						
						if(r instanceof XWPFHyperlinkRun) {
							XWPFHyperlink link = ((XWPFHyperlinkRun) r).getHyperlink(document);
							
							if(setup) {
								XWPFHyperlink tempLink = link;
								setup = false;
							}
							if(link != null && link.getURL().contains(replacement) && tempLink != link) {
								System.out.println(link.getURL());
								tempLink = link;
								linkCounter++;
							}
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
									
									if(r instanceof XWPFHyperlinkRun) {
										XWPFHyperlink link = ((XWPFHyperlinkRun) r).getHyperlink(document);
										
										if(setup) {
											XWPFHyperlink tempLink = link;
											setup = false;
										}
										if(link != null && link.getURL().contains(replacement) && tempLink != link) {
											//System.out.println(link.getURL());
											tempLink = link;
											linkCounter++;
										}
									}	
								}
							}
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
		
		return linkCounter;
		
	}//end of find_LinkDocX
}//end of FindFile
