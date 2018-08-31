package main.utils;

import java.io.File;
import java.io.FileFilter;



public class CommonUtil {

	public File[] getFiles(String folderPath) {
		File folder = new File(folderPath);
		File[] listOfFiles = folder.listFiles(new FileFilter() {          
		    public boolean accept(File file) {
		        return file.isFile();
		    }
		});
		return listOfFiles;
	}
	
	public String removeSpecialChar(String colname) {
		colname =  colname.toLowerCase().replaceAll("[^a-zA-Z0-9]+","_");
		return colname.startsWith("_")? colname.replace("_", "") : colname;
	}
}
