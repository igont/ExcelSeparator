import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Objects;
import java.util.Scanner;

public class Main
{
	public static Scanner sc = new Scanner(System.in);
	
	public static void main(String[] args) throws IOException
	{
		System.out.println("Archicad to dachanotes.ru cut translator");
		System.out.println("Creator: Igor Gontarenko");
		System.out.println("--------------------------------------------");
		
		
		File dir = new File(".");
		File[] filesList = dir.listFiles();
		
		assert filesList != null;
		for(File file : filesList)
		{
			if(file.isFile() && getFileExtension(file.getName()).equalsIgnoreCase("xlsx"))
			{
				System.out.println();
				System.out.println("Calculate file "+ file.getName());
				System.out.println("--------------------------------------------");
				calculate(file);
			}
		}
		sc.nextLine();
	}
	
	private static String getFileExtension(String fileName)
	{
		if(fileName == null || fileName.equals("")) return "undefined";
		int dotIndex = fileName.lastIndexOf(".");
		return (dotIndex == -1) ? "" : fileName.substring(dotIndex + 1);
	}
	private static String getFileName(String fileName)
	{
		if(fileName == null || fileName.equals("")) return "undefined";
		int dotIndex = fileName.lastIndexOf(".");
		return (dotIndex == -1) ? "" : fileName.substring(0, dotIndex) + " ";
	}
	
	private static void calculate(File file) throws IOException
	{
		ExcelParser excelParser = new ExcelParser(file);
		excelParser.setSheet(0);
		
		Estimate estimate = new Estimate();
		int row = 2;
		
		while(true)
		{
			String size;
			try
			{
				size = getFileName(file.getName()) + excelParser.readCell(row, 0);
			}
			catch(NullPointerException e)
			{
				break;
			}
			
			if(!Objects.equals(size, getFileName(file.getName())))
			{
				String length = excelParser.readCell(row, 1);
				String count = excelParser.readCell(row, 2);
				String comment = excelParser.readCell(row, 3);
				
				Pos pos = new Pos(length, count, comment);
				estimate.addPos(size, pos);
			}
			row++;
		}
		excelParser.close();
		write(estimate);
	}
	
	private static void write(Estimate estimate)
	{
		System.out.println();
		for(int index = 0; index < estimate.estimate.size(); index++)
		{
			Size size = estimate.estimate.get(index);
			String mat = size.name;
			
			File dir = new File(".");
			String path = dir.getAbsolutePath();
			path = path.substring(0, path.length() - 1) + mat + ".txt";
			
			try(FileWriter writer = new FileWriter(path, false))
			{
				for(Pos pos : size.positions)
				{
					String length = pos.length;
					String count = pos.count;
					String comment = pos.comment;
					
					String out = String.join(";", length, count, comment) + "\n";
					writer.append(out);
					writer.flush();
				}
			}
			catch(IOException e)
			{
				throw new RuntimeException(e);
			}
			
			System.out.println("File " + mat + " created");
		}
		System.out.println();
		System.out.println("All done!");
	}
}
