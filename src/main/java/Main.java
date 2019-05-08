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

	/**
	 * Description: Runs program to find and replace items in docx file
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws IOException{
		
		for (int i = 0; i < 3; i++) {
			System.out.print("--------------------------------------");
		}
		System.out.println();
		System.out.println("Finding Link with Port number in Directory " + DIRECTORY);
		
		File directory = new File(DIRECTORY);
		File[] fileList = directory.listFiles();
		
		for (int i = 0; i < fileList.length; i++) {
			if (fileList[i].isFile()) {
				tempPath = DIRECTORY + "\\" + fileList[i].getName();
				if(FindFile.find_LinkDocX(tempPath, PORT_REMOVAL) > 0) {
					System.out.println(fileList[i].getName());
				}
			}
		} 
		
	}
	
}