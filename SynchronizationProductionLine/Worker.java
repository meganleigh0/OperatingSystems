/*
 *Megan Griffin
 *2/22/2022
 *COSC 423
 */
public class Worker extends Thread {
    // Variables 
    final private boolean producer; // Boolean to distinguish prodcuer only/WorkerA 
    private boolean workerD = false; // Boolean to distinguish consumer only/WorkerD
    final private int limit = 24; // Widget production capacity  
    private int widgetCounter = 1; // Widget prodcution counter
    private ConveyorBelt beltIn;
    final private ConveyorBelt beltOut;
    private Widget widget;
    private int widgetWidgetID;

    // WorkerA constructor   
    public Worker(String name, boolean newProducer, ConveyorBelt b) {
        super(name);
        producer = newProducer;
        beltOut = b;
    }

    // WorkerBC constructor
    public Worker(String name, boolean newProducer, ConveyorBelt in, ConveyorBelt out) {
        super(name);
        producer = newProducer;
        beltIn = in;
        beltOut = out;
    }

    // WorkerD constructor
    public Worker(String name, boolean newProducer, ConveyorBelt in, ConveyorBelt out, boolean complete) {
        super(name);
        producer = newProducer;
        beltIn = in;
        beltOut = out;
        workerD = complete;
    }

    public void run() {
        // Runs until limit capacity is reached
        while(widgetCounter <= limit) {
            // Random call to sleep so workers handle widgets at varying rates
            ConveyorBelt.napping();

            // Condition only for WorkerA
            if(producer) {
                // Create new widget
                widget = new Widget(widgetCounter);
                widgetWidgetID = widget.getWidgetID();
                System.out.println(getName() + " is producing widget" + widgetCounter);
                int workers = widget.getNumWorkers();
                String handlers = widget.handled(workers);
                System.out.println(getName() + " is working on widget" + widgetCounter+handlers);
                beltOut.enter(widget, getName(), widgetWidgetID);
            } 
            
            else {
                // Remove the first widget from the inbound belt
                widget = beltIn.remove(getName());
                int workers = widget.getNumWorkers();
                String handlers = widget.handled(workers);
                System.out.println(getName() + " is retrieving widget" + widget.getWidgetID()+ handlers + " from the belt");

                // Add the wiget to the outbound belt 
                widget.workUpon();
                produce();

                // Condition only for WorkerD  
                if(workerD) {
                    beltOut.fullyProcessed(widget,getName(),widget.widgetID);
                } else {
                    beltOut.enter(widget,getName(), widget.getWidgetID());
                }
            }
            // Increment the widget counter 
            widgetCounter++;
        }
    }
    // Worker processing widget 
    public void produce() {
        int workers = widget.getNumWorkers();
        String handlers = widget.handled(workers);
        System.out.println(getName() + " is working on widget" + widget.getWidgetID()+handlers);
        ConveyorBelt.napping();
    }
}