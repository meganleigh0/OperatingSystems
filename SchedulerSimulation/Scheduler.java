package scheduler;

/**
 * <p>Title: Scheduler</p>
 * <p>Description: </p>
 * <p>Copyright: Copyright (c) 2015, 2004 by Matt Evett</p>
 * @author Matt Evett
 * @version 2.0
 * Manages the scheduling for the OS.  You'll need to extend this to make a concrete class.
 * It should be possible to make various concrete classes that implement different kinds of
 * schedulers (round robin, etc.)
 */

public abstract class Scheduler {
	protected volatile Job currentlyRunningJob; // current job

	public synchronized Job getRunningJob() {
		return currentlyRunningJob;
	}


	/**
	 * Add new job to readyQ.  This should interract with blockTilThereIsAJob appropriately.
	 * This method will probably be executed by the Submittor thread, while 
	 * blockTilThereIsAJob will probably be executed by the kernel (SystemSimulator) thread.
	 */
	public abstract void add( Job J );

	/**
	 * Remove job from readyQ.  No need to synchronize, as no thread blocks on "full" buffer
	 */
	public abstract void remove( Job J );

	/**
	 * @return true if there are any Jobs in the scheduler (could be running, ready, or 
	 * waiting on IO [in later version of project])
	 */
	public synchronized boolean hasJobs() {
		return hasJobsQueued() || hasRunningJob();
	}

	/**
	 * return true if the scheduler has jobs in any queues (ready, or IO queues 
	 * [in later version of project]
	 */
	public abstract boolean hasJobsQueued();

	/**
	 * Return true if there is a Job currently running
	 */
	public synchronized boolean hasRunningJob()
	{
		return( null != currentlyRunningJob);
	}

	public synchronized void clearRunningJob() {
		currentlyRunningJob = null;
	}	

	/**
	 * If the ready queue is empty, return false.
	 * Otherwise, start the next job in the queue, returning true.  If the queue is empty
	 * return false.
	 * Make the next job in the ready queue run. You should probably
	 * invoke Thread.start() on it.
	 */
	public abstract boolean makeRun();


	/**
	 * blockTilThereIsAJob()  Invoked by OS simulator when it wants to get a new Job to
	 * run.  This won't actually start any Job just yet, rather this will block the thread
	 * running this code (the OS simulator) until there is something on the ready queue.
	 * 
	 * Your implementation might be based on a remove() method of a SharedBuffer class in
	 * the Producer-Consumer problem.  The OS is acting like a consumer, trying to acquire
	 * a new Job from a potentially empty queue.  The "producer" will be the Submittor
	 * thread.
	 * 
	 */
	public abstract  void  blockTilThereIsAJob(); 
}