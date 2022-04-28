import java.io.*;
import java.util.Scanner;
import java.util.ArrayList;
import java.util.Collections;
/*
 * Megan Griffin
 * 4/22/2022
 * COSC 423 Computer Operating Systems
 * Disk Scheduler Assignment 
 */

public class DiskScheduler {
    // S : the number of milliseconds elapsed since the submission of the previous request
    // T : the track number requested
    private int s;
    private int t;
    public static void main(String [] args) {
        // Variables 
        int [][] reqsInt = getFileData();
        Requests[] reqs = new Requests[reqsInt.length];

        for(int i = 0; i < reqs.length; i++) {
            reqs[i] = new Requests(reqsInt[i][0], reqsInt[i][1]);
        }
        // F.C.F.S. - First Come First Serve
        FCFS fcfs = new FCFS (reqs);
        fcfs.printReq();
        System.out.println("TOTAL TIME: " + fcfs.getTime() + "\n");

        // LOOK
        LOOK look = new LOOK (reqs);
        look.sort();
        look.printReq();
        System.out.println("TOTAL TIME: " + look.getTime());

        // S.S.T.F - Shortest Seek Time First
        SSTF sstf = new SSTF (reqs);
        sstf.sort();
        sstf.printReq();
        System.out.println("TOTAL TIME: " + sstf.getTime() + "\n");

    }

    public static int[][] getFileData (){
        // Input file
        File fileName = new File("Input.txt");
        try {
            // Count all values in file  
            Scanner sc = new Scanner(fileName);
            int aCount = 0;
            while (sc.hasNextInt()) {
                aCount ++;
                sc.nextInt();
            }
            sc.close();
            int rCount = 0;
            // First values in the file = number of tracks
            // Each row contains two integers
            rCount = (aCount - 1)/2;
            int [][] reqs = new int [rCount][2];
            int numTracks; 
            Scanner readReq = new Scanner(fileName);
            numTracks = readReq.nextInt();
            for(int i = 0; i < rCount; i++) {
                for(int j = 0; j <2; j++) {
                    reqs[i][j] = readReq.nextInt();
                }
            }
            readReq.close(); 
            return reqs;        
        }
        catch (IOException ex){ 
            System.out.println("Input File, " + fileName + ", not found.");
            return null;
        } 
    }
}

class Requests {
    // S : the number of milliseconds elapsed since the submission of the previous request
    // T : the track number requested
    // Variables
    private int s;
    private int t;
    public Requests (int s, int t) {
        this.s = s;
        this.t = t;
    }
    // Get methods
    public int getS () {
        return s;
    }

    public int getT () {
        return t;
    }
}

class FCFS {
    // Variables
    private Requests [] reqs; 
    private int time; 

    public FCFS(Requests [] reqs) {
        this.reqs = reqs;
        this.time = 0;
    }

    public int getTime () {
        return time;
    }

    public void incTime (int reqTime) {
        time = time + reqTime;
    }

    public void printReq () {
        System.out.println("====================FCFS====================");
        for(int i = 0; i <reqs.length; i++) {
            incTime(reqs[i].getS());
            int moreTime;
            if (i != 0 ) {
                moreTime = Math.abs(reqs[i-1].getT() - reqs[i].getT());
                incTime(moreTime);
            } 
            System.out.println("[TIME: " + time + "] SERVICING track: " + reqs[i].getT());
        }
        System.out.println();
    }
}

class LOOK {
    // Variables
    private Requests [] reqs; 
    private int time;
    private ArrayList<Requests> low;
    private ArrayList<Requests> up;

    public LOOK(Requests [] reqs) {
        this.reqs = reqs;
        this.time = 0;
        this.low = new ArrayList<Requests>(reqs.length);
        this.up = new ArrayList<Requests>(reqs.length);
    }

    public int getTime () {
        return time;
    }

    public void incTime (int reqTime) {
        time = time + reqTime;
    }
    // Divide array in upper and lower half
    public void split () {
        for (int i = 1; i< reqs.length-1; i++) {
            if (reqs[i].getT() < reqs[0].getT()) {
                low.add(reqs[i]);
            }
            else {
                up.add(reqs[i]);
            }
        }
    }
    // Sort lower half of the array
    public void sortLow () {
        for (int i = 0; i < low.size()-1; i++){
            int index = i;
            for (int j = i+1; j < low.size(); j++) {
                if (low.get(i).getT() > low.get(index).getT()) {
                    index = j;
                }
            }
            Collections.swap(low, index, i);
        }
    }
    // Sort upper half of the array
    public void sortUp () {
        for (int i = 0; i < up.size()-1; i++){
            int index = i;
            for (int j = i+1; j < up.size(); j++) {
                if (up.get(i).getT() < up.get(index).getT()) {
                    index = j;
                }
            }
            Collections.swap(up, index, i);
        }
    }

    public void sort () {
        split();
        sortLow();
        sortUp();
    }

    public void printReq () {
        System.out.println("====================LOOK====================");
        System.out.println("[TIME: " + time + "] SERVICING track: " + reqs[0].getT());
        int moreTime;
        for(int i = 0; i <low.size(); i++) {
            incTime(low.get(i).getS());
            if (i != 0 ) {
                moreTime = Math.abs(low.get(i-1).getT() - low.get(i).getT());
                incTime(moreTime);
            } 
            System.out.println("[TIME: " + time + "] SERVICING track: " + low.get(i).getT());
        }
        for(int i = 0; i <up.size(); i++) {
            incTime(up.get(i).getS());
            if (i != 0 ) {
                moreTime = Math.abs(up.get(i-1).getT() - up.get(i).getT());
                incTime(moreTime);
            } 
            System.out.println("[TIME: " + time + "] SERVICING track: " + up.get(i).getT());
        }
        System.out.println();
    }
}

class SSTF {
    // Variables
    private Requests [] reqs; 
    private int time; 
    public SSTF(Requests [] reqs) {
        this.reqs = reqs;
        this.time = 0;
    }

    public int getTime () {
        return time;
    }

    public void incTime (int reqTime) {
        this.time = this.time + reqTime;
    }

    public void sort () {
        for (int i = 1; i< reqs.length-1; i++) {
            int minInd = i;
            int diff = Integer.MAX_VALUE;
            for(int j = i; j<reqs.length; j++) {
                int compared = Math.abs(reqs[i-1].getT() - reqs[j].getT());
                if (compared < diff) {
                    minInd = j;
                    diff = compared;
                }
            }
            Requests temp = reqs[minInd];
            reqs[minInd] = reqs[i];
            reqs[i] = temp;
        }
    }

    public void printReq () {
        System.out.println("====================SSTF====================");
        for(int i = 0; i <reqs.length; i++) {
            incTime(reqs[i].getS());
            int moreTime;
            if (i != 0 ) {
                moreTime = Math.abs(reqs[i-1].getT() - reqs[i].getT());
                incTime(moreTime);
            } 
            System.out.println("[TIME: " + time + "] SERVICING track: " + reqs[i].getT());
        }
        System.out.println();
    }

}
