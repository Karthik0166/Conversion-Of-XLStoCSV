package main;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.commons.lang.StringUtils;

public class Read {
	
	public void read()
	{
		try {
			FileInputStream fileInputStream = new FileInputStream("D:/work/csv/NEWGS-DRA_GS_DSR_wc1il_SP-NP_v8.11 R5.0.1+G8.xls");
			HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
			HSSFSheet worksheet = workbook.getSheet("Connections");			
			List list1=new ArrayList();
			int rowCount=2;
			boolean emptyRow=false;
			int totalRows=0;
			for(int i=0;i<worksheet.getLastRowNum()-1;i++)
			{
				HSSFRow row1 = worksheet.getRow(rowCount);
//				HSSFCell cell1 = row1.getCell(2);
				if(cellToString(row1.getCell(2))==null||cellToString(row1.getCell(2)).isEmpty())
				{
					emptyRow=true;
					list1.add("empty");
					totalRows++;					
				}
				else
				{
					emptyRow=false;
					list1.add(cellToString(row1.getCell(2)));
				}
				if(emptyRow)
				{
					list1.add("emptyRow");
				}
				else
				{
//					HSSFCell cell2 = row1.getCell(13);
					if(cellToString(row1.getCell(13))==null||cellToString(row1.getCell(13)).isEmpty())
					{
						list1.add("emptyCol");
					}
					else
					{
						if(cellToString(row1.getCell(2)).contains("IP"))
						{
							String ip=cellToString(row1.getCell(13));
							if(ip.contains("0000:0000:0000"))
							{
								ip=ip.replaceAll("0000:0000:0000", "");
							}
							else if (ip.contains("0000:0000"))
							{
								ip=ip.replaceAll("0000:0000", "");
							}
							if(ip.contains(":000"))
							{
								ip=ip.replaceAll(":000", ":");
							}
							if(ip.contains(":00"))
							{
								ip=ip.replaceAll(":00", ":");
							}
							if(ip.contains(":00"))
							{
								ip=ip.replaceAll(":00", ":");
							}
							list1.add(ip);
						}
						else if(cellToString(row1.getCell(13)).equals("TCP"))
						{
							String tcp=cellToString(row1.getCell(13)).replaceAll("TCP", "Tcp");
							list1.add(tcp);
						}
						else if(cellToString(row1.getCell(13)).equals("IP Address"))
						{
							String ipadd=cellToString(row1.getCell(13)).replaceAll("IP Address", "Ip");
							list1.add(ipadd);
						}
						else if(cellToString(row1.getCell(13)).contains("FQDN"))
						{
							String fqdn=cellToString(row1.getCell(13));
							String[] stFqdn=fqdn.split("\\s+");
							String actual="";
							for(int j=0;j<stFqdn.length;j++)
							{
								actual=actual+toProperCase(stFqdn[j]);
							}

							list1.add(actual);
						}
						else if(cellToString(row1.getCell(2)).contains("FQDN"))
						{
							String value=cellToString(row1.getCell(13));
							String[] ipcheck=value.split(":");
							int flag=list1.size();
							if(ipcheck.length>5)
							{
								String ip=cellToString(row1.getCell(13));
								if(ip.contains("0000:0000:0000"))
								{
									ip=ip.replaceAll("0000:0000:0000", "");
								}
								else if (ip.contains("0000:0000"))
								{
									ip=ip.replaceAll("0000:0000", "");
								}
								if(ip.contains(":000"))
								{
									ip=ip.replaceAll(":000", ":");
								}
								if(ip.contains(":00"))
								{
									ip=ip.replaceAll(":00", ":");
								}
								if(ip.contains(":00"))
								{
									ip=ip.replaceAll(":00", ":");
								}
								list1.set(flag-4, ip);
								list1.add("emptyCol");
							}
							else
							{
								list1.add(value);
							}
						}
						else
						{
							list1.add(cellToString(row1.getCell(13)));
						}
						
					}
				}
				rowCount++;
				
			}			
			Write write=new Write();
			write.write(list1,totalRows);
					
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	static String toProperCase(String s) {
	    return s.substring(0, 1).toUpperCase() +
	               s.substring(1).toLowerCase();
	}
	public static String cellToString(HSSFCell cell) {  
	    int type;
	    Object result = null;
	    type = cell.getCellType();

	    switch (type) {

	        case Cell.CELL_TYPE_NUMERIC: // numeric value in Excel
	        case Cell.CELL_TYPE_FORMULA: // precomputed value based on formula
	            result = cell.getNumericCellValue();
	            break;
	        case Cell.CELL_TYPE_STRING: // String Value in Excel 
	            result = cell.getStringCellValue();
	            break;
	        case Cell.CELL_TYPE_BLANK:
	            result = "";
	        case Cell.CELL_TYPE_BOOLEAN: //boolean value 
	            result: cell.getBooleanCellValue();
	            break;
	        case Cell.CELL_TYPE_ERROR:
	        default:  
	            throw new RuntimeException("There is no support for this type of cell");                        
	    }

	    return result.toString();
	}
}
