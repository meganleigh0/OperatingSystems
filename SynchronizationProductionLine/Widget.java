/*
 *Megan Griffin
 *2/22/2022
 *COSC 423
 */
public class Widget {
    // Variables
    int widgetID;
    int numWorkers;

    // Method to increment widgetID with production count 
    public Widget(int productionCount) {
        widgetID = productionCount;
        numWorkers = 1;
    }

    // Method to log the amount of workers that have handled a widget
    public void workUpon() {
        numWorkers++;
    }

    // Method to get widgetID
    public int getWidgetID() {
        return widgetID;
    }

    //Method to return the number of workers
    public int getNumWorkers() {
        return numWorkers;
    }
    
    //Method to return a string of workers that have handled each widget
    public String handled(int workedUpon){
        int handle = workedUpon;
        if(handle == 1)
        {
            return(" <handled by A> ");
        }
        if(handle == 2)
        {
            return(" <handled by A,B> ");
        }
        if(handle == 3)
        {
            return(" <handled by A,B,C> ");
        }
        if(handle == 4)
        {
            return(" <handled by A,B,C,D> ");
        }
        return null;
    }
}