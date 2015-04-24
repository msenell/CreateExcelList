package mainPack;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.CellView;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class ExcelFunctions 
{
	static boolean chkSize, chkBorder;
	
	
	public ExcelFunctions() 
	{
		chkBorder = true;
		chkSize = true;
	}
	
	protected Sheet getSheet(File f, String path)
	{
		Sheet s1 = null;
		Workbook w1 = null;
		path = path +"\\" + f.getName();
		
		try {
			w1=Workbook.getWorkbook(new File(path));
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		s1 = w1.getSheet(0);
		return s1;
	}
	
	protected String[] getARow(Sheet s1)
	{
		String[] aRow = new String[7];
		aRow[0] = "";
		aRow[1] = s1.getCell(0,12).getContents().toString();
		aRow[2] = s1.getCell(1,8).getContents().toString();
		aRow[3] = s1.getCell(3,8).getContents().toString();
		aRow[4] = s1.getCell(1,27).getContents().toString();
		aRow[5] = s1.getCell(12,8).getContents().toString();
		aRow[6] = s1.getCell(10,5).getContents().toString();
		
		
		return aRow;
	}
	
	protected void write2Excel(WritableWorkbook wtWB, String[] aRow, int row)
	{
		WritableFont font = new 
				WritableFont(WritableFont.ARIAL,11,WritableFont.BOLD,false);
		if(!aRow[0].equals("No"))
		{
			aRow[0] = String.valueOf(row-1);
			font = new WritableFont(WritableFont.ARIAL,11,WritableFont.NO_BOLD,false);
		}
		WritableSheet wtSH;  //Çýktý Liste'nin Sheet Deðiþkeni

		wtSH = wtWB.getSheet(0);
		for(int col = 0; col<aRow.length; col++)
		{
			//System.out.println(aRow[col]);
			writeOnCell(wtSH, col, row, aRow[col], font);
		}
		
	}
	
	protected WritableWorkbook createWrite(File lst)
	{
		WritableWorkbook wtWB = null; //Çýktý Liste'nin Workbook Deðiþkeni
		WorkbookSettings GenelAyarlar=new WorkbookSettings();
		GenelAyarlar.setLocale(Locale.ENGLISH);
		
		
		try {
			wtWB = Workbook.createWorkbook(lst, GenelAyarlar); //Create WorkBook
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		wtWB.createSheet("Sayfa1", 0);
		
		return wtWB;
		
	}
	
	private void writeOnCell(WritableSheet wtSH, int SutunNo, int SatirNo, String GelenHucreDegeri, WritableFont font)
	{
		
		WritableCellFormat cellFormat = new WritableCellFormat(font);
		if(chkBorder)
		try {
			
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		} catch (WriteException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} 
		Label cell = new Label(SutunNo, SatirNo, GelenHucreDegeri, cellFormat);
		
		try {
			wtSH.addCell(cell);
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	protected void cellFormatting(WritableWorkbook wb)
	{
		WritableSheet sh = wb.getSheet(0);
		CellView cw = new CellView();
		for(int i = 0; i<7; i++)
		{
			cw = sh.getColumnView(i);
			cw.setAutosize(true);
			sh.setColumnView(i, cw);
		}
	}
}
