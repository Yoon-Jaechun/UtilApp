package com.utilApp.main;

import java.io.File;

public class Test {
	
	public void method(File file){
		
//		if(file.exists()){
//			File[] files = file.listFiles();
//			
//			for(File f : files){
//				if(f.isDirectory()){
//					
//				}
//			}
//		}else{
		if(file.exists() == true){
			return;
		}else{
			boolean r = file.mkdirs();
		}
	}
	
	public static void main(String args[]){
		Test t = new Test();
		File file = new File("c:\\test1\\test2\\test3");
		t.method(file);
	}
}
