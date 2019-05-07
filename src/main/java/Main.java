package main.java;

import java.io.IOException;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;

import main.java.ReadDocX.ReadDocX;
import main.java.SearchDocX.FindFile;

public class Main {
	private static final String PATH = 
			"C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files\\ACC105 FIG.docx";
	private static final String DIRECTORY = 
			"C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files";
	
	private static final String TEST = 
			"C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files\\Test Template.docx";

	public static void main(String[] args) throws IOException{
		
		//for (int i = 0; i < 4; i++) {
			//System.out.print("--------------------------------------");
		//}
		//System.out.println("Class ReadDocX, Reading " + PATH);
		//for (int i = 0; i < 4; i++) {
			//System.out.print("--------------------------------------");
		//}
		//ReadDocX.readDocXFile(PATH);
		
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		System.out.println("Class FindFile, Finding files " + DIRECTORY);
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		FindFile.findFile(DIRECTORY);
		
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		System.out.println("Total Files: " + FindFile.getFileTotal(DIRECTORY));
		
		//Test find and remove
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		FindFile.find_Remove_DocX(TEST, "<<age>>");
		
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		System.out.println("Total Files: " + FindFile.getFileTotal(DIRECTORY));
		
	}
	
}