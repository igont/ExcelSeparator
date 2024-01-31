import java.util.ArrayList;
import java.util.List;

public class Size
{
	public Size(String name, Pos position)
	{
		this.name = name;
		this.positions.add(position);
	}
	
	public String name;
	public List<Pos> positions = new ArrayList<>();
}
