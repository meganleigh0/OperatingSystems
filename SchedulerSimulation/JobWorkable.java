package scheduler;

/**
 * Implementations of this interface are meant to simulate the work of a process
 * in our simulated system.
 * The interface contains a single method, doWork().  Each Job contains an object
 * (its "work") that
 * implements this interface. When a Job is running it will invoke doWork() on 
 * its work about every 10msec.  Your implementation of doWork() should return in 
 * no more than a few milliseconds.  In my solution I just had doWork generate
 * a bit of output.
 * Because in the follow-up to this assignment Job will rely on being able to handle 
 * interrupts to implement time sharing,
 * your implementation of doWork should not invoke interrupted() or isInterrupted()
 * and it should not catch InterruptExceptions (or, if it does, it should rethrow
 * them).  
 * @author matt
 * @version 1.0
 * <p>Copyright: Copyright (c) 2015 by Matt Evett</p>
 *
 */
public interface JobWorkable {
	public void doWork();
}