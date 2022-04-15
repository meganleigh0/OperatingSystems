The scheduler simulation simulation program consists of:
A thread simulating the OS kernel as an instance of the SystemSimulator class.
A set of threads simulating the processes managed by the kernel. Each thread is an instance of the Job class, and simulates a single process.
Each Job contains an instance of a class implementing the JobWorkable interface. The execution of this object simulates the work done by the process it is simulating.
A factory that generates JobWorkables.
A thread simulating users periodically submitting new processes to the kernel to be executed (an instance of the Submittor class). 
An input file that controls how often a thread submits a new process to the kernel, and the duration of that process's simulated CPU burst.
