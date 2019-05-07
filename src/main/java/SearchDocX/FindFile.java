package main.java.SearchDocX;

import java.io.File;
import java.io.IOException;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;

public class FindFile {
	
	public static void findFile(String fileName, File directory) {
		File[] list = directory.listFiles();
		
		if(list != null) {
			for (File file : list) {
				if (file.isDirectory()) {
					findFile(fileName, file);
				}
				else if (fileName.equalsIgnoreCase(file.getName())) {
					System.out.println(file.getParentFile());
				}
			}
		}
	}
	
	public static void find_Replace_DocX() throws IOException, InvalidFormatException {
		
	}
}
