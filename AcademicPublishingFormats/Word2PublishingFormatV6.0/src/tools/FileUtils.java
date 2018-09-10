package tools;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

public class FileUtils {

	public static File createFile(String fileName){
		File convertedFile = null;
		try {
			/*
			 * create a file to save the converted latex stream.
			 * begin
			 */
			convertedFile = new File(fileName);
			if(convertedFile.exists()){
				LogUtils.logShort("The file \""+convertedFile.getAbsolutePath()+"\" exists.\nIt needs to be replaced.\n");
				convertedFile.delete();
//				LogUtils.logShort("The file \""+convertedFile.getAbsolutePath()+"\" is deleted.\n");
			}
			convertedFile.createNewFile();
			LogUtils.logShort("The file \""+convertedFile.getAbsolutePath()+"\" is created.\n");
			/*
			 * end
			 * create a file to save the converted latex stream.
			 */
		} catch (Exception e) {
			LogUtils.log(e.getMessage());
		}
		return convertedFile;
	}
	
	public static void writeFile(StringBuilder laTeXString, File file){
		if(laTeXString != null){
			BufferedWriter writer;
			try {
				writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file), "UTF-8"));
				writer.write(laTeXString.toString());
				writer.flush();
				writer.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				LogUtils.log(e.getMessage());
			}
		}
	}
	
	public static void copyFile(File originalFile, File convertedFile, String fileName){
		try {
			LogUtils.logShort("write new file...");
			String folderPath = convertedFile.getAbsolutePath();
			folderPath = folderPath.substring(0, folderPath.lastIndexOf(File.separator));
			InputStream is = new FileInputStream(originalFile);
			LogUtils.logShort(".");
			Files.copy(is, convertedFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
			LogUtils.logShort(".");
			is.close();
			LogUtils.logShort("copy finish\n");
		} catch (Exception e) {
			LogUtils.log(e.getMessage());
		}
	}
}
