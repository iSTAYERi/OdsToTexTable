import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;

import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;


public class MainOdsToTexTable {
	
	public static void readODS(File file){
		Sheet sheet;
		String outputStr;
		try {
            //Getting the 0th sheet for manipulation| pass sheet name as string
            sheet = SpreadSheet.createFromFile(file).getSheet(0);
             
            //Get row count and column count
            int nColCount = sheet.getColumnCount();
            int nRowCount = sheet.getRowCount();
            
            outputStr = "\\begin{center}\n"
            		+ "\\begin{small}\n"
            		+ "\\begin{longtable}{";
            for (int i = 0; i < nColCount; i++){
            	outputStr = outputStr + "|c";
            }
            outputStr = outputStr + "|}\n"
            		+ "\\hline\n";
            
            System.out.println("Rows :"+nRowCount);
            System.out.println("Cols :"+nColCount);
            //Iterating through each row of the selected sheet
            MutableCell cell = null;
            for(int nRowIndex = 0; nRowIndex < nRowCount; nRowIndex++){
            	//Iterating through each column
            	int nColIndex = 0;
            	for( ;nColIndex < nColCount; nColIndex++){
            		cell = sheet.getCellAt(nColIndex, nRowIndex);
            		System.out.print(cell.getValue() + " col " + " ");
            		if (nColIndex == nColCount - 1){
            			outputStr = outputStr + cell.getValue();
            		}
            		else {
						outputStr = outputStr + cell.getValue() + " & ";
					}
            		
            	}
            	System.out.println();
            	outputStr = outputStr + "\\\\\n\\hline\n";
             }
            
            outputStr = outputStr + "\\end{longtable}\n"
            		+ "\\end{small}\n"
            		+ "\\end{center}\n";
            
//            System.out.println("String output:\n" + outputStr);
            
            File outFile = new File(file.getPath() + ".tex");
            try{
                PrintWriter writer = new PrintWriter(outFile);
                writer.print(outputStr);
                writer.close();
            } catch (IOException e) {
               System.out.println(e);
            }
		} catch (IOException e) {
        	   e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		File file = new File("/home/stayer/Documents/testTable.ods");
		readODS(file);

	}

}
