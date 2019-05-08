package main.java;

import java.io.File;
import java.io.IOException;

import main.java.SearchDocX.FindFile;

//Courses we know for sure have links: EDU512

public class Main {
	private static final String DIRECTORY = 
			"C:\\Users\\CatherineKhongsaly\\Dropbox\\Project_Workplace\\WorkProjects\\Files";
	private static final String PORT_REMOVAL = ":2048";
	private static final String TEXT_REMOVAL = "http";
	private static String tempPath;

	public static void main(String[] args) throws IOException{
		
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		System.out.println("Finding Link with Port number");
		System.out.println(DIRECTORY);
		
		File directory = new File(DIRECTORY);
		File[] fileList = directory.listFiles();
		
		for (int i = 0; i < fileList.length; i++) {
			if (fileList[i].isFile()) {
				tempPath = DIRECTORY + "\\" + fileList[i].getName();
				System.out.println("Path: " + tempPath);
				System.out.println("Number of links: " + FindFile.find_LinkDocX(tempPath, PORT_REMOVAL));
				System.out.println("Number of text found in docx: " + FindFile.find_TextDocX(tempPath, TEXT_REMOVAL));
			}
		} 
		
	}
	
}