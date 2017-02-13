package com.stta.utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Read_XLS {
	
	public String filelocation;
	public FileInputStream ipstr=null;
	public FileOutputStream opstr=null;
	private HSSFWorkbook wb=null;
	private HSSFSheet ws=null;
	
	
	public Read_XLS(String filelocation){
		this.filelocation=filelocation;
		try{
			ipstr=new FileInputStream(filelocation);
			wb=new HSSFWorkbook(ipstr);
			ws=wb.getSheetAt(0);
			ipstr.close();
		}
		catch(Exception e){
			e.printStackTrace();
		}
		
		
	}
	
	//To retrive NO Of Rows from .xls files sheets
	public int retriveNoOfRows(String wsName){
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1)
			return 0;
		else{
			
			ws=wb.getSheetAt(sheetIndex);
			int rowCount=ws.getLastRowNum()+1;
			return rowCount;
		}
	//	return (Integer) null;
	}
	
	//To retrive No Of Columns from .xls file's sheets
	public int retrieveNoOfCols(String wsName){
		
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1){
			return 0;
			
		}
		else{
			ws=wb.getSheetAt(sheetIndex);
			int colCount=ws.getRow(sheetIndex).getLastCellNum();
			return colCount;
		}
		//return 0;
		
	}
	//To retrive SuiteToRun and CaseToRun flag of test suite and testCase
	public String retriveToRunFlag(String wsName,String colName,String rowName){
		int sheetIndex =wb.getSheetIndex(wsName);
		if(sheetIndex==-1)
			return null;
		else{
			int rowNum=retriveNoOfRows(wsName);
			int colNum=retrieveNoOfCols(wsName);
			int colNumber=-1;
			int rowNumber=-1;
			
			HSSFRow Suiterow=ws.getRow(0);
			for(int i=0; i<colNum;i++){
				if(Suiterow.getCell(i).getStringCellValue
						().equals(colName.trim())){
					colNumber=i;
				}
				
			}
			if(colNumber==-1){
				return "";
			}
			for (int j=0; j<rowNum;j++){
				HSSFRow Suitecol=ws.getRow(j);
				if(Suitecol.getCell(0).getStringCellValue().equals(rowName.trim())){
					rowNumber=j;
				}
				
			}
			
			if(rowNumber==-1)
				return "";
			HSSFRow row=ws.getRow(rowNumber);
			HSSFCell cell=row.getCell(colNum);
			if(cell==null)
				return "";
			String value=cellToString(cell);
			return value;
		}
		
		//return null;
	}

	private String cellToString(HSSFCell cell) {
		// TODO Auto-generated method stub
		return null;
	}

}
