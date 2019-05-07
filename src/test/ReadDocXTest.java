package test;

import static org.junit.Assert.assertNotNull;

import org.junit.Test;

import main.java.ReadDocX.ReadDocX;

public class ReadDocXTest {
	private static final String TEST = "C:\\Users\\eoi61\\Dropbox\\Project_Workplace\\WorkProjects\\Files\\Test Template.docx";
	
	@Test
	public void readDocXFile() {
		System.out.println("--------------------------------------");
		System.out.println("Class ReadDocX, Reading " + TEST);
		System.out.println("--------------------------------------");
		//Test to ensure that test file works
		
		ReadDocX.readDocXFile(TEST);
		
	}

}
