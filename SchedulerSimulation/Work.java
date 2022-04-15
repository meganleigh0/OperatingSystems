package scheduler;
/**
 * <p>Title: Work</p>
 * <p>Description: </p>
 * @author Megan Griffin
 */
public class Work implements JobWorkable{
	int jobCount;

	
	@Override
	public void doWork() {
		
		System.out.println(Thread.currentThread().getName()+" is completing its task #"+jobCount);
		jobCount++;
	}

}