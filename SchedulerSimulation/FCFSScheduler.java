package scheduler;
import java.util.concurrent.ConcurrentLinkedQueue;
/**
 * <p>Title: FCFSScheduler</p>
 * <p>Description: Component of the simulate operating system that encapsulates FCFS job scheduling.</p>
 * <p>Copyright: Copyright (c) 2015, 2004</p>
 * <p>Company: </p>
 * @author Matt Evett
 * @version 2.0
 * Modified by Megan Griffin
 */

public class FCFSScheduler extends Scheduler {
	//data structure to support FCFS scheduler
	ConcurrentLinkedQueue<Job> readyQ= new ConcurrentLinkedQueue<Job>();
 
  /**
   * If the ready queue is empty, return false.
   * Otherwise, start the next job in the queue, returning true.  If the queue is empty
   * return false.
   * Make the next job in the ready queue run. You should probably
   * invoke Thread.start() on it.
   */
  public boolean makeRun()
  {
	  Job currjob=readyQ.remove();
	  if(currjob!=null)
	  {
		  this.currentlyRunningJob=currjob;
		  currentlyRunningJob.start();
		  return true;
	  }
	  return false;
  }
  
  /**
   * blockTilThereIsAJob()  Invoked by OS simulator when it wants to get a new Job to
   * run.  Will block if the ready queue is empty until a Job is added to the queue.
   */
  public synchronized void  blockTilThereIsAJob() {
	  if (hasRunningJob()) 
		  return;
	  while(!hasJobsQueued())
	  {
		  System.out.println("blockingTilThere's a job");
	  		try {
				this.wait();
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	  }
	  System.out.println("evidently there is now a job on readyQ");
  }

@Override
public synchronized void add(Job J) {
	System.out.println(J.getName()+" notifying the ready readyQ");
	readyQ.add(J);
	this.notify();
	
}

@Override
public void remove(Job J) {
	readyQ.remove();
	
}

@Override
public boolean hasJobsQueued() {
	if (readyQ.size()>0)
		return true;
	return false;
}
}