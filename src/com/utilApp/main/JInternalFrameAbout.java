package com.utilApp.main;

import java.awt.EventQueue;

import javax.swing.JInternalFrame;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.plaf.basic.BasicInternalFrameTitlePane.SystemMenuBar;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class JInternalFrameAbout extends JInternalFrame {

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					JInternalFrameAbout frame = new JInternalFrameAbout();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public JInternalFrameAbout() {
		setMaximizable(true);
		setClosable(true);
		setTitle("about");
		setBounds(100, 100, 234, 280);
		getContentPane().setLayout(null);
		
		JButton btnClose = new JButton("Close");
		btnClose.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
			}
		});
		btnClose.setBounds(41, 10, 97, 23);
		getContentPane().add(btnClose);
		
		JLabel lblMadeByYoon = new JLabel("Since 2015.10.10\r\n");
		lblMadeByYoon.setBounds(41, 31, 163, 23);
		getContentPane().add(lblMadeByYoon);
		
		JLabel lblNewLabel = new JLabel("Version 1.06");
		lblNewLabel.setBounds(41, 73, 163, 15);
		getContentPane().add(lblNewLabel);
		
		JLabel lblNewLabel_1 = new JLabel("Build : jde-7u10");
		lblNewLabel_1.setBounds(41, 110, 179, 15);
		getContentPane().add(lblNewLabel_1);
		
		JLabel lblNewLabel_2 = new JLabel("Made by Yoon JaeChun");
		lblNewLabel_2.setBounds(41, 147, 163, 15);
		getContentPane().add(lblNewLabel_2);

	}
}
