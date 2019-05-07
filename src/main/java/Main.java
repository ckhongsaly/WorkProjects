package main.java;

import main.java.ReadDocX.ReadDocX;
import main.java.SearchDocX.FindFile;

public class Main {
	private static final String PATH = 
			"C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files\\ACC105 FIG.docx";
	private static final String DIRECTORY = 
			"C:\\\\\\\\Users\\\\\\\\eoi61\\\\\\\\Dropbox\\\\\\\\Project_Workplace\\\\\\\\WorkProjects\\\\\\\\Files";

	public static void main(String[] args) {
		
		//System.out.println("--------------------------------------");
		//System.out.println("Class ReadDocX, Reading " + PATH);
		//System.out.println("--------------------------------------");
		//ReadDocX.readDocXFile(PATH);
		
		System.out.println("--------------------------------------");
		System.out.println("Class FindFile, Finding files " + DIRECTORY);
		System.out.println("--------------------------------------");
		FindFile.findFile(DIRECTORY);
	}
	
}