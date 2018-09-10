package main;

import base.AppEnvironment;
import ui.MainFrame;

public class StartThread extends Thread {

	public StartThread() {
		// TODO Auto-generated constructor stub
	}

	public StartThread(Runnable target) {
		super(target);
		// TODO Auto-generated constructor stub
	}

	public StartThread(String name) {
		super(name);
		// TODO Auto-generated constructor stub
	}

	public StartThread(ThreadGroup group, Runnable target) {
		super(group, target);
		// TODO Auto-generated constructor stub
	}

	public StartThread(ThreadGroup group, String name) {
		super(group, name);
		// TODO Auto-generated constructor stub
	}

	public StartThread(Runnable target, String name) {
		super(target, name);
		// TODO Auto-generated constructor stub
	}

	public StartThread(ThreadGroup group, Runnable target, String name) {
		super(group, target, name);
		// TODO Auto-generated constructor stub
	}

	public StartThread(ThreadGroup group, Runnable target, String name, long stackSize) {
		super(group, target, name, stackSize);
		// TODO Auto-generated constructor stub
	}

	public static void main(String args[]) {
		AppEnvironment.getInstance().init();
		
		MainFrame main = new MainFrame();
		
//		main.setResizable(false);
		main.setTitle("Word2AcademicFormatTool");
		
		main.setVisible(true);
		
		AppEnvironment.getInstance().setGui(main);
//		EventQueue.invokeLater(new Runnable() {
//			
//			@Override
//			public void run() {
//				// TODO Auto-generated method stub
//				
//				
//				
//			}
//		});
		
	}
	
	
}
