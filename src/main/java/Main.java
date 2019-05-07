package main.java;

import main.java.ReadDocX.ReadDocX;

public class Main {
	private static final String PATH = "C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files\\ACC105 FIG.docx";

	public static void main(String[] args) {
		
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading " + PATH);
		System.out.println("--------------------------------------");
		ReadDocX.readDocXFile(PATH);
	}
	
}