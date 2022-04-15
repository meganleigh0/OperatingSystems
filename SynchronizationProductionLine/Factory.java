/*
 *Megan Griffin
 *2/22/2022
 *COSC 423
 */
public class Factory {
    public static void main(String[] args) {

        // Create conveyor belts
        ConveyorBelt beltAtoB = new ConveyorBelt(); //belt for passing widgets from Worker A to Worker B
        ConveyorBelt beltBtoC = new ConveyorBelt(); //belt passing widgets from Worker B to Worker C
        ConveyorBelt beltCtoD = new ConveyorBelt(); //belt passing widgets from Worker C to Worker D
        ConveyorBelt beltProcessed = new ConveyorBelt(); //belt to store all processed widgets

        // Create worker threads with assigned boolean controls and belts
        Worker workerA = new Worker("Worker A", true, beltAtoB);
        Worker workerB = new Worker("Worker B", false, beltAtoB, beltBtoC);
        Worker workerC = new Worker("Worker C", false, beltBtoC, beltCtoD);
        Worker workerD = new Worker("Worker D", false, beltCtoD, beltProcessed, true);

        workerA.start();
        workerB.start();
        workerC.start();
        workerD.start();

        try {
            workerA.join();
            workerB.join();
            workerC.join();
            workerD.join();
        } catch (Exception e) { }

        System.out.println("");
        System.out.println("Widget Processing Complete!");
    }
}