package main;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Write {

	public void write(List list1, int totalRows)
	{
		Writer writer = null;

		try {
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd HH:mm:ss");
			Date date = new Date();	
			String first="";
			String second="";
			boolean flag=true;
		    writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("D:/work/csv/NEWGS-DRA_GS_DSR_wc1il_SP-NP_v8.11 R5.0.1+G8.csv"), "utf-8"));
		    writer.write("####################################################################################\n");
		    writer.write("#\n");
		    writer.write("# Date/Time Generated: "+dateFormat.format(date)+" Etc/UTC\n");
		    writer.write("# Exported Object: Conn\n");
		    writer.write("# Number of Records:"+totalRows+"\n");
		    writer.write("####################################################################################\n");
		    writer.write("Application Name,Screen,");
		    for(int i=0;i<list1.size();i++){
		    	if(i==0||i%2==0)
		    	{
		    		if(list1.get(i).toString().equalsIgnoreCase("empty"))
			    	{
			    		i++;	
			    		if(flag)
			    		{
			    			writer.write(first+"\n");
			    			flag=false;
			    		}			    		
			    		writer.write("Diameter,Conn,"+second+"\n");
			    		first="";
			    		second="";
			    	}
			    	else
			    	{
			    		if(flag)
			    		{
				    		if(!first.equalsIgnoreCase(""))
				    		{
				    			first=first+",";
				    		}
				    		first=first+list1.get(i).toString();
			    		}
			    	}
		    		
		    	}
		    	else
		    	{
		    		if(!second.equalsIgnoreCase(""))
		    		{
		    			second=second+",";
		    		}
		    		if(!list1.get(i).toString().equalsIgnoreCase("emptyCol"))
			    	{
		    			second=second+list1.get(i).toString();
			    	}
		    		
		    	}
		    	
		    }
		} catch (IOException ex) {
		  System.out.println("Exception:"+ex.getMessage());
		} finally {
		   try {writer.close();} 
		   catch (Exception ex) {
			   System.out.println("Exception:"+ex.getMessage());
		   }
		}
	}
}
