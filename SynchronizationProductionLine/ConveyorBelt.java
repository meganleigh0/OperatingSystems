/*
 *Megan Griffin
 *2/22/2022
 *COSC 423
 */
import java.util.LinkedList;
import java.util.Random;
public class ConveyorBelt {
    // Variables
    public static final int NAP_TIME = 5;
    // Conveyor belt has a maximum capacity of 3 widgets
    private static final int beltSize = 3;
    private LinkedList<Widget> widgetItems = new LinkedList<>();
    
    // Method for random call to sleep
    public static void napping() {
        int sleepTime = (int) (NAP_TIME * Math.random() );
        try {
            Thread.sleep(sleepTime*250); }
        catch(InterruptedException e) { }
    }

    // Method to enter widgets on belt 
    public synchronized void enter(Widget widget, String worker, int widgetNumber) {
        // Check if belt is full
        while (widgetItems.size() == 3) {
            try {
                int workers = widget.getNumWorkers();
                String handlers = widget.handled(workers);
                System.out.println("WARNING: "+ worker +" is waiting to put widget" + widgetNumber+handlers+ " on the belt");
                wait();
            } catch (InterruptedException e) { }
        }

        // Add widget to belt list
        widgetItems.add(widget);
        int workers = widget.getNumWorkers();
        String handlers = widget.handled(workers);
        System.out.println(worker + " is placing widget" + widgetNumber + handlers+ " on the belt");

        notify();
    }
    //Method to remove widgets from belt list
    public synchronized Widget remove(String worker) {
        Widget widget;

        // Condtion for empty belt list
        while(widgetItems.size() == 0) {
            System.out.println("WARNING: "+worker + " is idle!");
            try {
                wait();
            } catch (InterruptedException e) { }
        }

        widget = widgetItems.removeFirst();
        notify();

        return widget;
    }
    //Method for adding fully processed widgets to the processed belt 
    public synchronized  void fullyProcessed(Widget widget, String worker, int widgetNumber) {
        widgetItems.add(widget);
    }
}