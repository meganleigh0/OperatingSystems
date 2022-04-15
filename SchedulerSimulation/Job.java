package scheduler;
import java.util.concurrent.locks.Condition;
/**
 * <p>Title: Job</p>
 * <p>Description: </p>
 * <p>Copyright: Copyright (c) 2016, 2015, 2004 by Matt Evett</p>
 * <p>Company: </p>
 * @author Matt Evett
 * @version 2.1
 * Each instance of this class represents a process running within the virtual system being simulated
 * by the entire project.
 */

public class Job extends Thread {

	  private final SystemSimulator myOS; // OS
	  private final Condition myCondition; // This is associated with the OS's single reentrant lock
	  							//In my solution each Job has its own Condition that the OS simulator
	  							// uses to start that Job.
	  							// When we introduce time-slicing, the Job can be use this to suspend
	  							// itself, passing control back to the OS simulator.
	  
	  private final int burstTime; // job burst time
	  private final String name; // name of job
	  
	  private volatile boolean shouldRun = false; // true if job should be running
	  private volatile long startTime; // relativeTime when Job first starts running
	  							// You'll need to use this with Gannt chart calculations
	  private final JobWorkable work;  // What you want your Job to do as its "work".
	  
	  /*
	   * The burstDescription consists of a single integer, y, which is the
	   * the CPU burst duration.  In a later version of this program we'll augment
	   * the descriptors to allow for a sequence of CPU and IO burst lengths.
	   */
	  public Job(String burstDescriptor, SystemSimulator s, String name, JobWorkable workToDo) {
	    // Initialize stuff
	    myOS = s;
	    myCondition = s.getSingleThreadMutex().newCondition();
	    
	    burstTime = Integer.parseInt( burstDescriptor );
	    work = workToDo;
	    
	    this.name = name;
	  }
	  
	  /*
	   * An accessor, returning the starting time of the job.
	   */
	  protected long getStartTime()
	  {
	    return( startTime );
	  }

	  /*
	   * An accessor, returning the CPU burst time of the job.
	   */
	  protected int getBurstTime()
	  {
	    return( burstTime );
	  }

	  public Condition getMyCondition() {
		  return myCondition;
	  }
	  
	  /*
	   * returns false when the Job's burst time has been exhausted.
	   */
	  synchronized protected boolean shouldRun()
	  {
	    return( shouldRun );
	  }
	  
	  /*
	   * will be invoked when the burst time has been exhausted.
	   * This is a simple method that sets a flag which will eventually be accessed via a call to shouldRun().
	   * I will use a timer object to call this method at the appropriate time.
	   */
	  synchronized void pleaseStop()
	  {
	    shouldRun = false;
	  }
	  
	  /*
	   * return name of the job. Note that you can choose to use the inherited Thread.getName, 
	   * but if so, make sure you use the "name" argument appropriately in the Job constructor, above.
	   */
	  synchronized String getNameOf()
	  {
	    return( name );
	  }
	  
	  /*
	   * Can do pretty much anything but must return after the CPU burst time has elapsed.
	   * (Note that the Job's run-clock should not start "ticking" until run() is invoked! 
	   */
	  public void run()
	  {
		  //Should block here until the OS blocks itself on this Job's Condition
		  myOS.getSingleThreadMutex().lock();
		  
		  startTime = System.currentTimeMillis();
		  while (System.currentTimeMillis()-startTime < burstTime) {// Not yet exhausted my run-time
			  work.doWork(); // This should return in only a few milliseconds
			  try {
				sleep(10);
			  } catch (InterruptedException e) {
				  System.out.println(""+name+" is interrupted, hopefully only by TimeSlicer");
				  e.printStackTrace();
			  }
		  }
			  
		  exit();  //exit needs to signal the Condition, and release the lock
	  }
	  
	  /*
	   * should be last instruction in run(). Exit should eventually invoke myOS.exit();
	   */
	  public void exit()
	  {
	    myOS.exit();
	  }
	  

	}