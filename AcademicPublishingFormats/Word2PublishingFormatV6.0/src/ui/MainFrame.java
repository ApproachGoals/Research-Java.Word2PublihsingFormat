package ui;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.GraphicsConfiguration;
import java.awt.HeadlessException;

import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.GroupLayout;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JLabel;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;
import javax.swing.text.DefaultCaret;

import base.AppEnvironment;
import base.StaticConstants;
import db.AcademicStylePool;
import ui.actions.UIMainFrameButtonActionListener;

public class MainFrame extends GUI {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JLabel label = null;
	private JTextField textfield = null;
	private JButton fileOpenButton = null;
	private JButton generateButton = null;
	private JRadioButton ieee = null;
	private JRadioButton acm = null;
	private JRadioButton springer = null;
	private JRadioButton elsevier1 = null;
//	private JRadioButton elsevier2 = null;
	private JTextArea textarea = null;
	private JCheckBox isActiveTestMode = null;
	private JProgressBar progressBar = null;
	
	private int fileType = 1;
	
	private MainFrame me;
	
	public MainFrame getInstance() {
		if(me == null){
			me = new MainFrame();
		}
		
		return me;
	}
	
	public MainFrame() throws HeadlessException {
		// TODO Auto-generated constructor stub
		UIPanel basePanel = new UIPanel();
		UIPanel panel = new UIPanel();
		UIPanel subPanel = new UIPanel();
		UIPanel progressPanel = new UIPanel();
		
		BoxLayout boxLayout = new BoxLayout(basePanel, BoxLayout.Y_AXIS);
		basePanel.setLayout(boxLayout);
		basePanel.add(panel);
		basePanel.add(subPanel);
		basePanel.add(progressPanel);
		
		/*
		 * panel
		 */
		GroupLayout layout = new GroupLayout(panel);
		panel.setLayout(layout);
		layout.setAutoCreateGaps(true);
		layout.setAutoCreateContainerGaps(true);
		
		label = new JLabel("source file");
		textfield = new JTextField();
		fileOpenButton = new JButton("open");
		generateButton = new JButton("generate");
		fileOpenButton.setPreferredSize(new Dimension(103, 31));
		
//		fileOpenButton.setHorizontalAlignment(SwingConstants.RIGHT);
//		generateButton.setHorizontalAlignment(SwingConstants.RIGHT);
		
		fileOpenButton.addActionListener(new UIMainFrameButtonActionListener(this));
		generateButton.addActionListener(new UIMainFrameButtonActionListener(this));
		
		ieee = new JRadioButton(StaticConstants.BUTTON_CONVERTFILE_IEEE_TEXT, true);
		acm = new JRadioButton(StaticConstants.BUTTON_CONVERTFILE_ACM_TEXT);
		springer = new JRadioButton(StaticConstants.BUTTON_CONVERTFILE_SPRINGER_TEXT);
		elsevier1 = new JRadioButton(StaticConstants.BUTTON_CONVERTFILE_ELSEVIER1_TEXT);
//		elsevier2 = new JRadioButton(StaticConstants.BUTTON_CONVERTFILE_ELSEVIER2_TEXT);
		
		isActiveTestMode = new JCheckBox(StaticConstants.BUTTON_CONVERTFILE_FONTNAME_STRICT);
		isActiveTestMode.setSelected(false);
//		isActiveTestMode.addActionListener(new UIMainFrameButtonActionListener(this));
//		isStyleNameStrict.addActionListener(new UIMainFrameButtonActionListener(this));
		
		layout.setHorizontalGroup(
				layout.createSequentialGroup()
					.addComponent(label)
					.addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
							.addComponent(textfield)
							.addGroup(layout.createSequentialGroup()
									.addComponent(ieee)
									.addComponent(acm)
									.addComponent(springer)
									.addComponent(elsevier1)
//									.addComponent(elsevier2)
									.addComponent(isActiveTestMode)
									)
							)
					.addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
							.addComponent(fileOpenButton)
							.addComponent(generateButton)
							)
					
		);
		layout.setVerticalGroup(
				layout.createSequentialGroup()
				.addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
							.addComponent(label)
							.addComponent(textfield)
							.addComponent(fileOpenButton)
				)
				.addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
							.addComponent(ieee)
							.addComponent(acm)
							.addComponent(springer)
							.addComponent(elsevier1)
//							.addComponent(elsevier2)
							.addComponent(isActiveTestMode)
							.addComponent(generateButton)
						)
		);
		
		layout.linkSize(SwingConstants.HORIZONTAL, fileOpenButton, generateButton);
		
		ButtonGroup radioGroup = new ButtonGroup();
		radioGroup.add(ieee);
		radioGroup.add(acm);
		radioGroup.add(springer);
		radioGroup.add(elsevier1);
//		radioGroup.add(elsevier2);
		
		/*
		 * subpanel
		 */
		subPanel.setLayout(new BorderLayout());
		subPanel.setBorder( new TitledBorder ( new EtchedBorder (), "Display Area" ) );
		
		textarea = new JTextArea(15, 20);
		textarea.setEditable(false); // set textArea non-editable
		textarea.setBackground(Color.BLACK);
		textarea.setForeground(Color.GREEN);
		textarea.setWrapStyleWord(true);
        textarea.setLineWrap(true);
		
		DefaultCaret caret = (DefaultCaret) textarea.getCaret();
		caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);
		textarea.setCaret(caret);
		
		JScrollPane scroll = new JScrollPane(textarea);
		scroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);
		subPanel.add(scroll, BorderLayout.CENTER);
		
		/*
		 * progressPanel
		 */
		progressBar = new JProgressBar();
		progressBar.setMinimum(0);
		progressBar.setMaximum(100);
		
		progressPanel.setLayout(new BorderLayout());
		progressPanel.add(progressBar, BorderLayout.CENTER);
		
		progressBar.setStringPainted(true);
//		Dimension dm = new Dimension(this.getPreferredSize().width, progressBar.getPreferredSize().height);
//		progressBar.setPreferredSize(dm);
		progressBar.setValue(0);
		
		AppEnvironment.getInstance().setStylePool(new AcademicStylePool());
		AppEnvironment.getInstance().getStylePool().loadDefinition();
		
		this.add(basePanel);
		this.pack();
		
	}

	public MainFrame(GraphicsConfiguration gc) {
		super(gc);
		// TODO Auto-generated constructor stub
	}

	public MainFrame(String title) throws HeadlessException {
		super(title);
		// TODO Auto-generated constructor stub
	}

	public MainFrame(String title, GraphicsConfiguration gc) {
		super(title, gc);
		// TODO Auto-generated constructor stub
	}

	public JLabel getLabel() {
		return label;
	}

	public void setLabel(JLabel label) {
		this.label = label;
	}

	public JTextField getTextfield() {
		return textfield;
	}

	public void setTextfield(JTextField textfield) {
		this.textfield = textfield;
	}

	public JButton getFileOpenButton() {
		return fileOpenButton;
	}

	public void setFileOpenButton(JButton fileOpenButton) {
		this.fileOpenButton = fileOpenButton;
	}

	public JButton getGenerateButton() {
		return generateButton;
	}

	public void setGenerateButton(JButton generateButton) {
		this.generateButton = generateButton;
	}

	public JRadioButton getIeee() {
		return ieee;
	}

	public void setIeee(JRadioButton ieee) {
		this.ieee = ieee;
	}

	public JRadioButton getAcm() {
		return acm;
	}

	public void setAcm(JRadioButton acm) {
		this.acm = acm;
	}

	public JRadioButton getSpringer() {
		return springer;
	}

	public void setSpringer(JRadioButton springer) {
		this.springer = springer;
	}

	public JRadioButton getElsevier1() {
		return elsevier1;
	}

	public void setElsevier1(JRadioButton elsevier1) {
		this.elsevier1 = elsevier1;
	}

//	public JRadioButton getElsevier2() {
//		return elsevier2;
//	}
//
//	public void setElsevier2(JRadioButton elsevier2) {
//		this.elsevier2 = elsevier2;
//	}

	public JTextArea getTextarea() {
		return textarea;
	}

	public void setTextarea(JTextArea textarea) {
		this.textarea = textarea;
	}

	public int getFileType() {
		return fileType;
	}

	public void setFileType(int fileType) {
		this.fileType = fileType;
	}

	public JCheckBox getIsActiveTestMode() {
		return isActiveTestMode;
	}

	public void setIsActiveTestMode(JCheckBox isActiveTestMode) {
		this.isActiveTestMode = isActiveTestMode;
	}

	public JProgressBar getProgressBar() {
		return progressBar;
	}

	public void setProgressBar(JProgressBar progressBar) {
		this.progressBar = progressBar;
	}
	
}
