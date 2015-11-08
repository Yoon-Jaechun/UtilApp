package com.utilApp.java;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelController {
	public ExcelController(){
		
	}
	
	public List createFolderList(File path, int indexSheet) throws Exception{
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<String> list = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row != null){
				Cell cell = row.getCell(0);
				if(cell != null){
					File file = new File(cell.getStringCellValue());
					boolean isCreateFolder = file.mkdirs();
					if(isCreateFolder == true){
						list.add("생성된 폴더 : " + file.getAbsolutePath());
					}else if(isCreateFolder == false && file.exists() == true){
						list.add("이미 존재하는  폴더 : " + file.getAbsolutePath());
					}else{
						list.add("생성하지 못한 폴더 : " + file.getAbsolutePath());
					}
				}
			}
		}
	    return list;
	}
	
	public List<String> getLoadExcelOneColumnList(File path, int indexSheet) throws Exception{
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<String> list = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row != null){
				Cell cell = row.getCell(0);
				if(cell != null){
					list.add(cell.getStringCellValue());
				}
			}
		}
	    return list;
	}
	
	public List<Object[]> getLoadExcelTwoColumnList(File path, int indexSheet) throws Exception{
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<Object[]> list = new ArrayList<Object[]>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			Object[] object = null;
			for(int j = 0; j < row.getLastCellNum(); j++){
				Cell oneCell = row.getCell(0) != null ? row.getCell(0) : null;
				Cell twoCell = row.getCell(1) != null ? row.getCell(1) : null;
				Cell threeCell = row.getCell(2) != null ? row.getCell(2) : null;
				Cell fourCell = row.getCell(3) != null ? row.getCell(3) : null;
				
				
				if(oneCell == null){
					object = new Object[]{"", twoCell.getStringCellValue()};
				}else if(twoCell == null){
					object = new Object[]{oneCell.getStringCellValue(), ""};
				}else{
					object = new Object[]{oneCell.getStringCellValue(), twoCell.getStringCellValue()};
				}
			}
			list.add(object);
		}
	    return list;
	}
	
	public List<Object[]> getLoadExcelFourColumnList(File path, int indexSheet) throws Exception{
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<Object[]> list = new ArrayList<Object[]>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			Object[] object = null;
			for(int j = 0; j < row.getLastCellNum(); j++){
				Cell oneCell = row.getCell(0) != null ? row.getCell(0) : null;
				Cell twoCell = row.getCell(1) != null ? row.getCell(1) : null;
				Cell threeCell = row.getCell(2) != null ? row.getCell(2) : null;
				Cell fourCell = row.getCell(3) != null ? row.getCell(3) : null;
				
				if(oneCell == null){
					object = new Object[]{"불러올 파일 없음", "사용할 수 없음", "사용할 수 없음", "사용할 수 없음"};
				}else if(oneCell != null && twoCell == null && threeCell != null){
					if(fourCell == null){
						object = new Object[]{oneCell.getStringCellValue(), "", threeCell.getStringCellValue(), ""};
					}else{
						object = new Object[]{oneCell.getStringCellValue(), "", threeCell.getStringCellValue(), fourCell.getStringCellValue()};
					}
				}else if(oneCell != null && twoCell != null && threeCell == null){
					if(fourCell == null){
						object = new Object[]{oneCell.getStringCellValue(), twoCell.getStringCellValue(), "", ""};
					}else{
						object = new Object[]{oneCell.getStringCellValue(), twoCell.getStringCellValue(), "", fourCell.getStringCellValue()};
					}
				}else if(oneCell != null && twoCell != null || threeCell != null){
					if(fourCell == null){
						object = new Object[]{oneCell.getStringCellValue(), twoCell.getStringCellValue(), threeCell.getStringCellValue(), ""};
					}else{
						object = new Object[]{oneCell.getStringCellValue(), twoCell.getStringCellValue(), threeCell.getStringCellValue(), fourCell.getStringCellValue()};
					}
				}else{
					if(fourCell == null){
						object = new Object[]{oneCell.getStringCellValue(), twoCell.getStringCellValue(), threeCell.getStringCellValue(), ""};
					}else{
						object = new Object[]{oneCell.getStringCellValue(), twoCell.getStringCellValue(), threeCell.getStringCellValue(), fourCell.getStringCellValue()};
					}
				}
			}
			list.add(object);
		}
	    return list;
	}
	
	public List fetchFolderList(File path, int indexSheet) throws Exception{
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<String> list = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row != null){
				Cell cell = row.getCell(0);
				if(cell != null){
					File file = new File(cell.getStringCellValue());
					if(file.exists() == true){
						if(file.isDirectory() == true){
							File[] files = file.listFiles();
							for(int j = 0; j < files.length; j++){
								if(files[j].isDirectory()){
									list.add(files[j].getAbsolutePath());
								}
							}
						}
					}
				}
			}
		}
	    return list;
	}
	
	public List fetchFileList(File path, int indexSheet) throws Exception{
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<String> list = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row != null){
				Cell cell = row.getCell(0);
				if(cell != null){
					File file = new File(cell.getStringCellValue());
					if(file.exists() == true){
						if(file.isDirectory() == true){
							File[] files = file.listFiles();
							for(int j = 0; j < files.length; j++){
								if(files[j].isFile()){
									list.add(files[j].getAbsolutePath());
								}
							}
						}
					}
				}
			}
		}
	    return list;
	}
	
	public List renameFolderList(File path, int indexSheet) throws Exception{
		
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<String> list = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row != null){
				Cell cellOne = row.getCell(0);
				Cell cellTwo = row.getCell(1);
				
				if(cellOne != null && cellTwo != null){
					File fileOne = new File(cellOne.getStringCellValue());
					File fileTwo = new File(cellTwo.getStringCellValue());
					
					if(fileOne.exists() == true){
						if(fileOne.isDirectory() == true){
							boolean result = fileOne.renameTo(fileTwo);
							if(result == true){
								list.add(fileOne.getAbsolutePath() + " 폴더에서 " + fileTwo.getAbsolutePath() + " 폴더로 변경 되었습니다.");
							}else if(result == false && fileTwo.isDirectory() == true){
								list.add(fileTwo.getAbsolutePath() + " 폴더가 이미 존재하여 " + fileOne.getAbsolutePath() + " 폴더에서 " + fileTwo.getAbsolutePath() + " 폴더로 변경 되지 않았습니다.");
							}else{
								list.add(fileOne.getAbsolutePath() + " 폴더에서 " + fileTwo.getAbsolutePath() + " 폴더로 변경 되지 않았습니다.");
							}
						}else if(fileOne.isFile() == true){
							list.add(fileOne.getAbsolutePath() + " 파일은 다른 기능으로 통해서 변경하세요.");
						}
					}else{
						list.add(fileOne.getAbsolutePath() + " 경로가 존재하지 않습니다.");
					}
				}
			}
		}
	    return list;
	}
	
public List renameFileList(File path, int indexSheet) throws Exception{
		
		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<String> list = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row != null){
				Cell cellOne = row.getCell(0) != null ? row.getCell(0) : null;
				Cell cellTwo = row.getCell(1) != null ? row.getCell(1) : null;
				
				if(cellOne != null && cellTwo != null){
					File fileOne = new File(cellOne.getStringCellValue());
					File fileTwo = new File(cellTwo.getStringCellValue());
					
					if(fileOne.exists() == true){
						if(fileOne.isFile() == true){
							boolean result = fileOne.renameTo(fileTwo);
							if(result == true){
								list.add(fileOne.getAbsolutePath() + " 파일에서 " + fileTwo.getAbsolutePath() + " 파일로 변경 되었습니다.");
							}else if(result == false && fileTwo.isFile() == true){
								list.add(fileTwo.getAbsolutePath() + " 파일이 이미 존재하여 " + fileOne.getAbsolutePath() + " 파일에서 " + fileTwo.getAbsolutePath() + " 파일로 변경 되지 않았습니다.");
							}else{
								list.add(fileOne.getAbsolutePath() + " 파일에서 " + fileTwo.getAbsolutePath() + " 파일로 변경 되지 않았습니다.");
							}
						}else if(fileOne.isDirectory() == true){
							list.add(fileOne.getAbsolutePath() + " 폴더는 다른 기능으로 통해서 변경하세요.");
						}
					}else{
						list.add(fileOne.getAbsolutePath() + " 경로가 존재하지 않습니다.");
					}
				}
			}
		}
	    return list;
	}

	public List replaceConextFilList(File path, int indexSheet) throws Exception{

		InputStream inp = new FileInputStream(path.getAbsolutePath());
		Workbook wb = WorkbookFactory.create(inp);

		List<String> list = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(indexSheet);
		for(int i = 1; i <= sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row != null){
				Cell cellOne = row.getCell(0) != null ? row.getCell(0) : null;
				Cell cellTwo = row.getCell(1) != null ? row.getCell(1) : null;
				Cell cellThree = row.getCell(2) != null ? row.getCell(2) : null;
				Cell cellFour = row.getCell(3) != null ? row.getCell(3) : null;

				File fileOne = null;
				if(cellOne != null){
					fileOne = new File(cellOne.getStringCellValue());
				}
				
				if(cellOne == null){
					list.add("치환하려는 파일이 존재하지 않습니다.");
				}else if(cellTwo == null){
					list.add(fileOne.getAbsolutePath() + " 치환 전 문자가 존재하지 않습니다.");
				}else if(cellThree == null){
					list.add(fileOne.getAbsolutePath() + " 치환 후 문자가 존재하지 않습니다.");
				}else if(cellOne != null && cellTwo != null && cellThree != null){
					if(fileOne.exists() == true){
						if(fileOne.isFile() == true){
							if(cellFour == null){
								String context = FileUtils.readFileToString(fileOne, "UTF-8");
								context = context.replaceAll(cellTwo.toString(), cellThree.toString());
								FileUtils.writeStringToFile(fileOne, context,"UTF-8");
								list.add(context + "\n" + "[" + cellTwo.toString() + "] -> [" + cellThree.toString() + "]" + "\n" + fileOne.getAbsolutePath() + "(UTF-8) 파일이 변경 되었습니다.\n");
							}else{
								String context = FileUtils.readFileToString(fileOne, cellFour.getStringCellValue());
								context = context.replaceAll(cellTwo.toString(), cellThree.toString());
								FileUtils.writeStringToFile(fileOne, context, cellFour.getStringCellValue());
								list.add(context + "\n" + "[" + cellTwo.toString() + "] -> [" + cellThree.toString() + "]" + "\n" + fileOne.getAbsolutePath() + "(" + cellFour.toString() + ") 파일이 변경 되었습니다.\n");
							}
						}
					}else{
						list.add(fileOne.getAbsolutePath() + " 파일이 존재하지 않습니다.");
					}
				}
			}
		}
	    return list;
	}
}
