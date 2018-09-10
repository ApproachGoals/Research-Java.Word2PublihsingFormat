package ui.actions;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import base.AppEnvironment;
import lib.AbstractPublishingFormatFactory;
import lib.FormatConverter;
import tools.FileUtils;
import tools.LogUtils;
import ui.MainFrame;

public class UIMainFrameButtonActionListener implements ActionListener {

	private MainFrame gui = null;
	
	@Override
	public void actionPerformed(ActionEvent event) {
		// TODO Auto-generated method stub
		if(event.getSource().equals(gui.getGenerateButton())){
			String comment = gui.getTextarea().getText()+"\n===You have chosen to convert the file===\n";
			LogUtils.log(comment);
			
			if(gui.getIsActiveTestMode().isSelected()){
				AppEnvironment.getInstance().setActiveTestMode(true);
			} else {
				AppEnvironment.getInstance().setActiveTestMode(false);
			}
			
			AppEnvironment.getInstance();
			if(AppEnvironment.getCurrentFile()==null){
				comment = gui.getTextarea().getText()+"Error, convert process was not executed, because you have not specified the original word file.\n";
				LogUtils.log(comment);
			} else {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setDialogTitle("Save the converted file as");
				String jFileDialogFilterString = ""; 
				String jFileDialogFilterTip = "";
				if(gui.getFileType()==1){
					jFileDialogFilterString = "*.docx|Word Files";
					jFileDialogFilterTip = "docx";
				} else if(gui.getFileType()==2){
					jFileDialogFilterString = "*.tex|LaTeX Files";
					jFileDialogFilterTip = "tex";
				} else if(gui.getFileType()==3){
					jFileDialogFilterString = "*.pdf|PDF Files";
					jFileDialogFilterTip = "pdf";
				}
				FileNameExtensionFilter filter = new FileNameExtensionFilter(jFileDialogFilterString, jFileDialogFilterTip);
				fileChooser.setFileFilter(filter);

				int selectedValue = fileChooser.showSaveDialog(null);
				
				AbstractPublishingFormatFactory formatFactory = null;
				
				if(selectedValue == JFileChooser.APPROVE_OPTION){
					String fileName = fileChooser.getSelectedFile().getAbsolutePath();
					if(gui.getFileType()==1){
						if(fileName.endsWith(".docx")){
							
						} else {
							fileName += ".docx";
						}
					} else if(gui.getFileType()==2){
						if(fileName.endsWith(".tex")){
							
						} else {
							fileName += ".tex";
						}
					} else if(gui.getFileType()==3){
						//TODO PDF
					}
					
					File convertedFile = null;
					
					if(gui.getIeee().isSelected()){
						if(gui.getFileType()==1){
							if(fileName.indexOf("IEEE")>=0){
								
							} else {
								fileName = fileName.substring(0, fileName.length()-5)+"_IEEE"+fileName.substring(fileName.length()-5, fileName.length());
							}
							formatFactory = FormatConverter.getFactory(FormatConverter.FILE_FORMAT_MODE_OOXML_IEEE);
						} else if(gui.getFileType()==2){

						} else if(gui.getFileType()==3){
							//TODO pdf
						}
						
					} else if(gui.getAcm().isSelected()){
						if(gui.getFileType()==1){
							if(fileName.indexOf("ACM")>=0){
								
							} else {
								fileName = fileName.substring(0, fileName.length()-5)+"_ACM"+fileName.substring(fileName.length()-5, fileName.length());
							}
							formatFactory = FormatConverter.getFactory(FormatConverter.FILE_FORMAT_MODE_OOXML_ACM);
						} else if(gui.getFileType()==2){
							fileName = fileName.substring(0, fileName.length()-4)+"_ACM"+fileName.substring(fileName.length()-4, fileName.length());
						} else if(gui.getFileType()==3){
							//TODO pdf
						}
						
					} else if(gui.getSpringer().isSelected()) {
						if(gui.getFileType()==1){
							if(fileName.indexOf("SPRINGER")>=0){
								
							} else {
								fileName = fileName.substring(0, fileName.length()-5)+"_SPRINGER"+fileName.substring(fileName.length()-5, fileName.length());
							}
							formatFactory = FormatConverter.getFactory(FormatConverter.FILE_FORMAT_MODE_OOXML_SPRINGER);
						} else if(gui.getFileType()==2){
//							fileName = fileName.substring(0, fileName.length()-4)+"_SPRINGER"+fileName.substring(fileName.length()-4, fileName.length());
						} else if(gui.getFileType()==3){
							//TODO pdf
						}
					} else if(gui.getElsevier1().isSelected()) {
						if(gui.getFileType()==1){
							if(fileName.indexOf("ELSEVIER_1")>=0){
								
							} else {
								fileName = fileName.substring(0, fileName.length()-5)+"_ELSEVIER_1"+fileName.substring(fileName.length()-5, fileName.length());
							}
							formatFactory = FormatConverter.getFactory(FormatConverter.FILE_FORMAT_MODE_OOXML_ELSEVIER1);
						}
//					} else if(gui.getElsevier2().isSelected()) {
//						if(gui.getFileType()==1){
//							if(fileName.indexOf("ELSEVIER_2")>=0){
//								
//							} else {
//								fileName = fileName.substring(0, fileName.length()-5)+"_ELSEVIER_2"+fileName.substring(fileName.length()-5, fileName.length());
//							}
//							formatFactory = FormatConverter.getFactory(FormatConverter.FILE_FORMAT_MODE_OOXML_ELSEVIER2);
//						}
					} else {
						LogUtils.log("Sorry, the goal format was not recognized.");
					}

					if(gui.getIsActiveTestMode().isSelected()){
						AppEnvironment.getInstance().setActiveTestMode(true);
					} else {
						AppEnvironment.getInstance().setActiveTestMode(false);
					}
					
					convertedFile = FileUtils.createFile(fileName);
					FileUtils.copyFile(AppEnvironment.getCurrentFile(), convertedFile, null);
					LogUtils.log("File's context has been copied.");
					
					formatFactory.convertProcess(convertedFile);
					
					LogUtils.log("finish, file save under the path: "+convertedFile.getAbsolutePath());
					
				} else {
					LogUtils.log("Operation was cancelled.");
				}
			}
			
		} else if(event.getSource().equals(gui.getFileOpenButton())) {
			LogUtils.clear();
			LogUtils.log("===You have chosen to open the file===");
			
			JFileChooser fileChooser = new JFileChooser();
			fileChooser.setDialogTitle("Please choose the original word file");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("*.docx|MS-Word 2007 or later", "docx");
			fileChooser.setFileFilter(filter);
			
			int selectedValue = fileChooser.showOpenDialog(null);
			if(selectedValue == JFileChooser.APPROVE_OPTION){
				File file = fileChooser.getSelectedFile();
				if(file != null){
					gui.getTextfield().setText(file.getAbsolutePath());
					LogUtils.log("The file \""+file.getAbsolutePath()+"\" was chosen.");
					AppEnvironment.setCurrentFile(file);
				} else {
					LogUtils.log("The file doesn't exist.");
				}
			}
		} else {
			LogUtils.logShort("\n===Error, unkown control stream, an unkown signal was captured as button event.===\n");
		}
	}

	public UIMainFrameButtonActionListener(){
		
	}
	
	public UIMainFrameButtonActionListener(MainFrame gui){
		this.gui = gui;
	}
}
