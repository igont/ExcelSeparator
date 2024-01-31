import java.util.ArrayList;
import java.util.List;

public class Estimate
{
	public List<Size> estimate = new ArrayList<>();
	
	public void addPos(String sizeStr, Pos pos)
	{
		boolean find = false;
		for(Size size : estimate)
		{
			if(sizeStr.equals(size.name))
			{
				size.positions.add(pos);
				find = true;
				break;
			}
		}
		if(!find)
		{
			estimate.add(new Size(sizeStr, pos));
		}
	}
}
