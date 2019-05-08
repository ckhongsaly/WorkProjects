package main.java;

import java.io.File;
import java.io.IOException;

import main.java.SearchDocX.FindFile;

//Courses we know for sure have links: EDU512

public class Main {
	private static final String DIRECTORY = 
			"C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files";
	private static final String PORT_REMOVAL = ":0248";
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
				System.out.println("Number of links: " + FindFile.findLinkDocX(tempPath, PORT_REMOVAL));
			}
		} 
		
	}
	
}