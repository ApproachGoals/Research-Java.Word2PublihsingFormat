package ui.actions;

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.AbstractButton;
import javax.swing.BoxLayout;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;

import base.AppEnvironment;
import base.StaticConstants;
import tools.LogUtils;

public class UIMainFrameMenuActionListener implements ActionListener {

	@Override
	public void actionPerformed(ActionEvent event) {
		// TODO Auto-generated method stub
		if(event.getActionCommand().equals(StaticConstants.COMMAND_EXIT)){
			System.exit(0);
		} else if(event.getActionCommand().equals(StaticConstants.COMMAND_RULE_ORIGINALFILE)){
			JFileChooser fileChooser = new JFileChooser();
			fileChooser.setDialogTitle("Please choose the original word file");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("*.json", "json");
			fileChooser.setFileFilter(filter);
			
			int selectedValue = fileChooser.showOpenDialog(null);
			if(selectedValue == JFileChooser.APPROVE_OPTION){
				File file = fileChooser.getSelectedFile();
				if(file != null){
					AppEnvironment.getInstance().setNewStyleFile(file);
					AppEnvironment.getInstance().getStylePool().loadCDefinition();
				} else {
					LogUtils.log("The file doesn't exist.");
				}
			}
		} else if(event.getActionCommand().equals(StaticConstants.COMMAND_ABOUT)){
			
			JLabel label1= new JLabel("Paluno");
			label1.setAlignmentX(Component.CENTER_ALIGNMENT);
			JLabel label2= new JLabel("The Ruhr Institute for Software Technology");
			label2.setAlignmentX(Component.CENTER_ALIGNMENT);
			JLabel label3= new JLabel("Institut für Informatik und Wirtschaftsinformatik");
			label3.setAlignmentX(Component.CENTER_ALIGNMENT);
			JLabel label4= new JLabel("Universität Duisburg-Essen");
			label4.setAlignmentX(Component.CENTER_ALIGNMENT);
			JLabel label5= new JLabel("Author: Yuan Shi");
			label5.setAlignmentX(Component.CENTER_ALIGNMENT);
			
			JPanel panel = new JPanel();
			panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
			panel.setAlignmentX(Component.CENTER_ALIGNMENT);
			panel.add(label1);
			panel.add(label2);
			panel.add(label3);
			panel.add(label4);
			panel.add(new JLabel("\n"));
			panel.add(label5);
			JOptionPane.showMessageDialog(AppEnvironment.getInstance().getGui(), panel, "About the tool", JOptionPane.PLAIN_MESSAGE, null);
			
		} else {
			
			AbstractButton aButton = (AbstractButton) event.getSource();
//	        boolean selected = aButton.getModel().isSelected();
	        if(aButton.getText().equals("Word")){
	        	AppEnvironment.getInstance().getGui().setFileType(1);
	        } else if(aButton.getText().equals("LaTeX")){
	        	AppEnvironment.getInstance().getGui().setFileType(2);
	        } else if(aButton.getText().equals("PDF")){
	        	AppEnvironment.getInstance().getGui().setFileType(3);
	        } else {
	        	
	        }
		}
		
	}

	public UIMainFrameMenuActionListener() {
		// TODO Auto-generated constructor stub
		
	}
}
