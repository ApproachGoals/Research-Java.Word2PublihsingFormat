package tools;

import java.lang.reflect.InvocationTargetException;

import javax.swing.SwingUtilities;

import base.AppEnvironment;

public class LogUtils {

//	private static String comment;
	public static void clear(){
		try {
//			new Thread(new BackendThread(comment)).start();
			new Thread(new Runnable() {
				
				@Override
				public void run() {
					// TODO Auto-generated method stub
					try {
						SwingUtilities.invokeLater(new Runnable() {

							@Override
							public void run() {
								// TODO Auto-generated method stub
								AppEnvironment.getInstance().getGui().getTextarea().setText("");
							}
							
						});
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					
				}
			}).start();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void logShort(Object o) {
		if (AppEnvironment.getInstance()!=null && AppEnvironment.getInstance().getGui() != null) {
			
//			BackendThread logBackend = new BackendThread(comment);
//			Thread logThread = new Thread(logBackend);
//			logThread.start();
//			try {
//				SwingUtilities.invokeLater(new BackendThread(comment));
//			} catch (Exception e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
			try {
//				new Thread(new BackendThread(comment)).start();
				new Thread(new Runnable() {
					
					@Override
					public void run() {
						// TODO Auto-generated method stub
						try {
							SwingUtilities.invokeAndWait(new Runnable() {

								@Override
								public void run() {
									// TODO Auto-generated method stub
									String comment = AppEnvironment.getInstance().getGui().getTextarea().getText();
									if (o != null) {
										comment += o.toString();
									} else {
										comment += "unknown error.";
									}
									AppEnvironment.getInstance().getGui().getTextarea().setText(comment);
									System.out.println(o.toString());
								}
								
							});
						} catch (InvocationTargetException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						
					}
				}).start();
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			System.out.print(o.toString());
		}
	}

	public static void log(Object o) {
		if (AppEnvironment.getInstance()!=null && AppEnvironment.getInstance().getGui() != null) {
			
//			BackendThread logBackend = new BackendThread(comment);
//			Thread logThread = new Thread(logBackend);
//			logThread.start();
			try {
//				new Thread(new BackendThread(comment)).start();
				new Thread(new Runnable() {
					
					@Override
					public void run() {
						// TODO Auto-generated method stub
						try {
							SwingUtilities.invokeLater(new Runnable() {

								@Override
								public void run() {
									// TODO Auto-generated method stub
									String comment = AppEnvironment.getInstance().getGui().getTextarea().getText();
									if (o != null) {
										comment += o.toString() + "\n";
									} else {
										comment += "unknown error.\n";
									}
									AppEnvironment.getInstance().getGui().getTextarea().setText(comment);
									if(o!=null){
										System.out.println(o.toString());
									} else {
										System.out.println("message is null");
									}
								}
								
							});
						} catch (Exception e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						
					}
				}).start();
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
//			System.out.println(o.toString());
		}
	}
}
