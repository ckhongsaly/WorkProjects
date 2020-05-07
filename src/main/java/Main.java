package main.java;

import java.io.File;
import java.io.IOException;

import main.java.DocXToXML.DocXToXML;
import main.java.SearchDocX.FindFile;

//Courses we know for sure have links: EDU512

public class Main {

	//enter path
	private static final String DIRECTORY = 
			//"C:\\Users\\CatherineKhongsaly\\Dropbox\\Project_Workplace\\WorkProjects\\Files";
			//"C:\\Users\\CatherineKhongsaly\\Documents\\FIG\\SMU";
			"C:\\Users\\CatherineKhongsaly\\Documents\\Test Docs\\Week 1";
	private static final String PORT_REMOVAL = "dipity";
	private static String tempPath;

	/**
	 * Description: Runs program to find and replace items in docx file
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws IOException{
		
		File directory = new File(DIRECTORY);
		File[] fileList = directory.listFiles();
		int totalQuestion = 0;

		//This is for link replacement
		//int totalLink = 0;
		//int totalText = 0;
		
		//Gather List
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();

		int counter = 0;
		
		for (int i = 0; i < fileList.length; i++) {
		
			if (fileList[i].isFile()) {
				tempPath = DIRECTORY + "\\" + fileList[i].getName();
				
				//if(FindFile.find_LinkDocX(tempPath, PORT_REMOVAL) > 0) {
					//System.out.println(fileList[i].getName() + 
							//"|| Total Links: " + (FindFile.find_LinkDocX(tempPath, PORT_REMOVAL)));
					//totalText = FindFile.find_TextDocX(tempPath, PORT_REMOVAL) + totalText;
					//totalLink = FindFile.find_LinkDocX(tempPath, PORT_REMOVAL) + totalLink;
					//counter++;
				//}
				totalQuestion = DocXToXML.find_Questions(tempPath);

			}
		}

		

		System.out.println("Total number of files in directory: " + FindFile.get_File_Total(DIRECTORY));
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();

		System.out.println("Total number of questions in directory: " + totalQuestion);
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();


		//Start of remove link

		//System.out.println("Number of Files with link: " + counter);
		//for (int i = 0; i < 3; i++) {
			//System.out.print("--------------------------------------");
		//}
		//System.out.println();
		//System.out.println("Total embed links: " + totalLink);
		//for (int i = 0; i < 3; i++) {
			//System.out.print("--------------------------------------");
		//}
		//System.out.println();

		//System.out.println("Total text links: " + totalText);
		//for (int i = 0; i < 3; i++) {
			//System.out.print("--------------------------------------");
		//}

		/** Link Removal and Replacement */
		/**
		System.out.println();
		System.out.println("Link Replacement in Directory " + DIRECTORY);
		
		for (int i = 0; i < fileList.length; i++) {
			
			if (fileList[i].isFile()) {
				tempPath = DIRECTORY + "\\" + fileList[i].getName();
				
				if(FindFile.find_LinkDocX(tempPath, PORT_REMOVAL) > 0) {
					FindFile.find_RemoveLink_DocX(tempPath, PORT_REMOVAL, "");
				}
			}
		} 
		
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		*/
	}
	
}