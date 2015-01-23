package main;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Read {
	
	public void read()
	{
		try {
			FileInputStream fileInputStream = new FileInputStream("D:/work/Input.xls");
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
				if(row1.getCell(2).getStringCellValue()==null||row1.getCell(2).getStringCellValue().isEmpty())
				{
					emptyRow=true;
					list1.add("empty");
					totalRows++;
				}
				else
				{
					emptyRow=false;
					list1.add(row1.getCell(2).getStringCellValue());
				}
				if(emptyRow)
				{
					list1.add("emptyRow");
				}
				else
				{
//					HSSFCell cell2 = row1.getCell(13);
					if(row1.getCell(13).getStringCellValue()==null||row1.getCell(13).getStringCellValue().isEmpty())
					{
						list1.add("emptyCol");
					}
					else
					{
						list1.add(row1.getCell(13).getStringCellValue());
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
}
