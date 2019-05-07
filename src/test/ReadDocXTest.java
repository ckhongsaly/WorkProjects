package test;

import org.junit.Test;

import main.java.ReadDocX.ReadDocX;

public class ReadDocXTest {
	private static final String TEST = "C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files\\Test Template.docx";
	
	//Test needs to be done better, does not actually test program but run it
	@Test
	public void readDocXFileTest() {
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading " + TEST);
		System.out.println("--------------------------------------");
		//Test to ensure that test file works
		
		ReadDocX.readDocXFile(TEST);
	}
	
	//Test needs to be done better, does not actually test program but run it
	@Test
	public void readHeaderFooterDocXTest() {
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading Header " + TEST);
		System.out.println("--------------------------------------");
		ReadDocX.readHeaderFooterDocX(TEST);
	}
	
	//Test needs to be done better, does not actually test program but run it
	@Test
	public void readParagraphTest() {
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading Paragraph " + TEST);
		System.out.println("--------------------------------------");
		ReadDocX.readParagraphDocX(TEST);
	}
		
	//Test needs to be done better, does not actually test program but run it
	@Test
	public void readTableDocxTest() {
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading Table " + TEST);
		System.out.println("--------------------------------------");
		ReadDocX.readTableDocX(TEST);
	}
		
	//Test needs to be done better, does not actually test program but run it
	@Test
	public void readStyleTest() {
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading Style " + TEST);
		System.out.println("--------------------------------------");
		ReadDocX.readStyle(TEST);
	}
		
	//Test needs to be done better, does not actually test program but run it
	@Test
	public void readImageTest() {
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading Image " + TEST);
		System.out.println("--------------------------------------");
		ReadDocX.readImage(TEST);
	}

}
