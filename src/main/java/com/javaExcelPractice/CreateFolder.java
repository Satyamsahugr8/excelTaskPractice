package com.javaExcelPractice;

import java.io.File;

public class CreateFolder {
	
	public static void main(String[] args) {
		
		File file = new File("temp");
		
		if(file.mkdir()) {
			System.out.println("Folder created");
		} else {
			System.out.println("No");
		}
	}

}
