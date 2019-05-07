package main.java;

import main.java.ReadDocX.ReadDocX;

public class Main {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println("Class ReadDocx.");
		String path = "C:/Users/CatherineKhongsaly/Documents/Test/Test Template.docx";
		
		ReadDocX reader = new ReadDocX();
		reader.readDocXFile(path);
	}
	
}