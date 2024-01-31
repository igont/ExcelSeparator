import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelParser
{
	private final XSSFWorkbook book;
	private XSSFSheet sheet;
	private final File file;
	
	public ExcelParser(File file) throws IOException
	{
		this.file = file;
		book = new XSSFWorkbook(new FileInputStream(file));
		setSheet(0);
	}
	
	public void setSheet(int sheetNum)
	{
		this.sheet = book.getSheetAt(sheetNum);
	}
	
	private Cell getCell(int row, int column)
	{
		try
		{
			return sheet.getRow(row).getCell(column);
		}
		catch(NullPointerException e)
		{
			return null;
		}
	}
	
	public String readCell(int row, int column)
	{
		Cell cell = getCell(row, column);
		return readDifferentCellTypes(cell, cell.getCellType());
	}
	
	private String readDifferentCellTypes(Cell cell, CellType cellType)
	{
		switch(cellType)
		{
			case _NONE, BLANK ->
			{
				return "";
			}
			case NUMERIC ->
			{
				return cell.getNumericCellValue() + "";
			}
			case STRING ->
			{
				return cell.getStringCellValue();
			}
			case FORMULA ->
			{
				CellType cachedFormulaResultType = cell.getCachedFormulaResultType();
				return readDifferentCellTypes(cell, cachedFormulaResultType);
			}
			case BOOLEAN ->
			{
				return cell.getBooleanCellValue() + "";
			}
			case ERROR ->
			{
				return "Error";
			}
			default -> throw new IllegalStateException("Unexpected value: " + cellType);
		}
	}
	
	public XSSFSheet getSheet()
	{
		return sheet;
	}
	
	public void close()
	{
		try
		{
			//book.write(new FileOutputStream(file));
			book.close();
		}
		catch(IOException e)
		{
			throw new RuntimeException(e);
		}
	}
}
