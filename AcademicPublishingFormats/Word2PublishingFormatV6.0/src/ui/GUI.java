package ui;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.GraphicsConfiguration;
import java.awt.HeadlessException;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.net.URL;
import java.util.Set;

import javax.swing.ButtonGroup;
import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JRadioButtonMenuItem;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import base.StaticConstants;
import ui.actions.UIMainFrameMenuActionListener;

public class GUI extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public GUI() throws HeadlessException {
		// TODO Auto-generated constructor stub
		int fSize = 12;
		
		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException ex) {
            ex.printStackTrace();
        }
		if(dim.getWidth()>1990){
			setDefaultSize(24);
			fSize = 24;
		} else if(dim.getWidth()>1700) {
			setDefaultSize(18);
			fSize = 18;
		} else if(dim.getWidth()>1400) {
			setDefaultSize(16);
			fSize = 16;
		} else {
			setDefaultSize(12);
			fSize = 12;
		}
		
		int w = Math.max(dim.width/2, 600);
		int y = Math.max(dim.height/2, 455);
		this.setPreferredSize(new Dimension(w, y));
		this.setMinimumSize(new Dimension(w, y));
		this.setSize(w, y);
		
		this.setLocation(dim.width/2-this.getSize().width/2, dim.height/2-this.getSize().height/2);
		try {
			URL iconURL = getClass().getResource("/sources/myude_icon.png");
			ImageIcon icon = new ImageIcon(iconURL);
			this.setIconImage(icon.getImage());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		UIMainFrameMenuActionListener al = new UIMainFrameMenuActionListener();
		
		JMenuBar menuBar = new JMenuBar();
		
		JMenu menu = new JMenu("File");
		Font menuF = new Font(menu.getFont().getName(), menu.getFont().getStyle(), fSize);
		menu.setFont(menuF);
		
		JMenuItem menuItem = new JMenuItem();
		menuItem.setText(StaticConstants.COMMAND_RULE_ORIGINALFILE);
		menuItem.addActionListener(al);
		menuItem.setMnemonic(KeyEvent.VK_F);
		Font font = new Font(menuItem.getFont().getName(), menuItem.getFont().getStyle(), fSize);
		menuItem.setFont(font);
		menu.add(menuItem);
		
		
		menuItem = new JMenuItem("About");
		menuItem.addActionListener(al);
		menuItem.setMnemonic(KeyEvent.VK_A);
		menuItem.setFont(font);
		menu.add(menuItem);
		
		menuItem = new JMenuItem("Exit");
		menuItem.addActionListener(al);
		menuItem.setMnemonic(KeyEvent.VK_Q);
		menuItem.setFont(font);
		menu.add(menuItem);
		
		menuBar.add(menu);
		

		menu = new JMenu("File Formats");
		menu.setFont(menuF);
		ButtonGroup group = new ButtonGroup();
		
		JRadioButtonMenuItem formatItem = new JRadioButtonMenuItem("Word");
		formatItem.addActionListener(al);
		formatItem.setSelected(true);
		group.add(formatItem);
		menu.add(formatItem);
		menuItem.setFont(font);
		formatItem = new JRadioButtonMenuItem("LaTeX");
		formatItem.addActionListener(al);
		formatItem.setSelected(false);
		formatItem.setEnabled(false);
		group.add(formatItem);
		menu.add(formatItem);
		menuItem.setFont(font);
		formatItem = new JRadioButtonMenuItem("PDF");
		formatItem.addActionListener(al);
		formatItem.setSelected(false);
		formatItem.setEnabled(false);
		group.add(formatItem);
		menu.add(formatItem);
		menuItem.setFont(font);
		
		menuBar.add(menu);
		
		this.setJMenuBar(menuBar);
		
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
	}

	public GUI(GraphicsConfiguration gc) {
		super(gc);
		// TODO Auto-generated constructor stub
	}

	public GUI(String title) throws HeadlessException {
		super(title);
		// TODO Auto-generated constructor stub
	}

	public GUI(String title, GraphicsConfiguration gc) {
		super(title, gc);
		// TODO Auto-generated constructor stub
	}

	public void setDefaultSize(int size) {

        Set<Object> keySet = UIManager.getLookAndFeelDefaults().keySet();
        Object[] keys = keySet.toArray(new Object[keySet.size()]);

        for (Object key : keys) {

            if (key != null && key.toString().toLowerCase().contains("font")) {

                Font font = UIManager.getDefaults().getFont(key);
                if (font != null) {
                    font = font.deriveFont((float)size);
                    UIManager.put(key, font);
//                    System.out.println(key);
                }
            }

        }

    }
}
