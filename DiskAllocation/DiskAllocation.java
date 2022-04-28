/*
 * Megan Griffin
 * 4/22/2022
 * COSC 423 Computer Operating Systems
 * Disk Space Allocation Extra Credit Assignment
 */

import java.io.File;
import java.io.FileNotFoundException;
import java.util.Scanner;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class DiskAllocation {
    public static void main(String[] args) {
        String inputFile = "disk.dat";
        int totBlock = 0;
        String lineOfData;
        String cmd = "";

        // Read data from input file 
        File file = new File(inputFile);
        Scanner scan = null;
        try {
            scan = new Scanner(file);
        } catch (FileNotFoundException e) {
            System.out.println(inputFile + " Not Found");
            System.exit(1);
        }
        totBlock = Integer.parseInt(scan.next());

        // Contiguous Allocation
        System.out.println("=============== CONTIGUOUS ALLOCATION ===============");
        DiskSimulator sd = new DiskSimulator(totBlock);
        System.out.println("totBlock = " + totBlock);
        scan.nextLine();
        while(scan.hasNextLine()){
            lineOfData = scan.nextLine();
            String[] tokens = lineOfData.split("\"");
            for(int i=0; i < tokens.length; i++) {
                tokens[i] = tokens[i].trim();
            }
            cmd = tokens[0];
            if(cmd.equals("add") || cmd.equals("del") || cmd.equals("print") || cmd.equals("read")) {
                switch(cmd) {
                    case "add":
                    sd.contiguousAllocation(tokens[1], Integer.parseInt(tokens[2]));
                    break;
                    case "del":
                    sd.deallocate(tokens[1]);
                    break;
                    case "print":
                    sd.printDir();
                    break;
                    case "read":
                    sd.read(tokens[1]);
                    break;
                }
            } else {
                System.out.println("Command is not invalid. Please verify input file data.");
            }

        }
        scan.close();
        System.out.println("=============== CONTIGUOUS ALLOCATION STATISTICS ===============");
        sd.printStats();
        System.out.println();
        System.out.println("=============== END OF CONTIGUOUS ALLOCATION ===============");

        // Indexed Allocation
        file = new File(inputFile);
        scan = null;
        try {
            scan = new Scanner(file);
        } catch (FileNotFoundException e) {
            System.out.println("FileNotFoundException: " + e);
        }
        System.out.println("=============== INDEXED ALLOCATION ===============");
        DiskSimulator sd2 = new DiskSimulator(totBlock);
        System.out.println("totBlock = " + totBlock);
        scan.nextLine();
        while(scan.hasNextLine()){
            lineOfData = scan.nextLine();
            String[] tokens = lineOfData.split("\"");
            for(int i=0; i < tokens.length; i++) {
                tokens[i] = tokens[i].trim();
            }
            cmd = tokens[0];
            if(cmd.equals("add") || cmd.equals("del") || cmd.equals("print") || cmd.equals("read")) {
                switch(cmd) {
                    case "add":
                    sd2.indexedAllocation(tokens[1], Integer.parseInt(tokens[2]));
                    break;
                    case "del":
                    sd2.deallocate(tokens[1]);
                    break;
                    case "print":
                    sd2.printDir();
                    break;
                    case "read":
                    sd2.read(tokens[1]);
                    break;
                }
            } else {
                System.out.println("Invalid cmd. Check input file.");
            }

        }
        scan.close();
        System.out.println("=============== INDEXED ALLOCATION STATISTICS ===============");
        sd2.printStats();
        System.out.println();
        System.out.println("=============== END OF INDEXED ALLOCATION ===============");
    }
}
class DiskSimulator {
    HashMap allocList = new HashMap(); // K = index start of file, V = size of file
    HashMap tempList = new HashMap(); // K = index start of hole, V = size of hole
    HashMap directory = new HashMap(); // K = index start of file, V = fileName
    HashMap adjustedDir = new HashMap<String, String>();
    HashMap fileNumbers = new HashMap(); // K = fileName, V = number
    ArrayList keySet = new ArrayList();
    ArrayList matched = new ArrayList();
    boolean flag = false;
    int[] detailsArray;
    int[] temp;
    int totalHoles;
    int fatBlocks;
    int number = 1;
    int index;
    int size;
    int segmentSize = 1;
    int block;
    int count;
    int notAlloc = 0;
    int numMoves;
    int totMoves = 0;

    public DiskSimulator(int s) {
        this.size = s;
    }

    public HashMap generateTempList() {
        tempList = new HashMap();
        segmentSize = 1;
        for(int i=0; i < size; i++) {
            if(allocList.containsKey(i)) {
                i += (Integer) allocList.get(i) - 1; // Go to end of file 
                segmentSize = 1;
            } else if (segmentSize == 1) {
                tempList.put(i, segmentSize);
                segmentSize++;
                index = i;
            }else {
                tempList.replace(index, segmentSize);
                segmentSize++;
            }
        }
        return tempList;
    }

    public void contiguousAllocation(String fileName, int sizeOfFile) {
        if(sizeOfFile <= 0) {
            System.out.println("Size of file cannot be less than 1!");
            return;
        }
        // add to allocated list
        if(allocList.isEmpty()) {
            allocList.put(0, sizeOfFile);
            directory.put(0, fileName);
            System.out.println("File " + fileName + " was added successfully");
            return;
        } else {
            tempList = generateTempList();
            // Contigous 
            if(contiguous(fileName, sizeOfFile))
                return;
        }
        notAlloc++;
        System.out.println("Not enough space! File " + fileName + " not added.");
    }

    public void indexedAllocation(String fileName, int sizeOfFile) {
        if(sizeOfFile <= 0) {
            System.out.println("Size of file cannot be less than 1!");
            return;
        }
        // add to allocated list
        if(allocList.isEmpty()) {
            // calculate FAT blocks
            sizeOfFile = calcFAT(sizeOfFile);
            allocList.put(0, sizeOfFile);
            directory.put(0, fileName);
            System.out.println("File " + fileName + " was added successfully");
            return;
        } else {
            tempList = generateTempList();
            // Indexed
            if(indexed(fileName, sizeOfFile))
                return;
        }
        notAlloc++;
        System.out.println("Not enough space! File " + fileName + " not added.");
    }

    public void sort(int arr[]) {
        // Function to sort array using insertion sort
        int n = arr.length;
        for (int i = 1; i < n; ++i) {
            int key = arr[i];
            int j = i - 1;

            while (j >= 0 && arr[j] > key) {
                arr[j + 1] = arr[j];
                j = j - 1;
            }
            arr[j + 1] = key;
        }
    }

    public void deallocate(String file) {
        boolean flag = false;
        for(int i=0; i < size; i++) {
            if(directory.get(i) != null)
                if(directory.get(i).equals(file)) {
                    allocList.remove(i);
                    directory.remove(i);
                    System.out.println("File " + file + " was deleted successfully");
                    return;
                }
        }
    }

    public void createAdjustedDirectory() {
        keySet = new ArrayList();
        adjustedDir = new HashMap();
        matched = new ArrayList();
        String blocks = "";

        // Referenced: https://stackoverflow.com/questions/10462819/get-keys-from-hashmap-in-java
        for ( Object key : allocList.keySet() ) {
            keySet.add(key);
        }

        blocks = "";
        number = 1;
        for(int i=0; i < keySet.size(); i++) {
            blocks = "";
            block = (Integer) keySet.get(i);
            for(int j=0; j < (Integer) allocList.get(keySet.get(i)); j++) {
                blocks += block + " ";
                block++;
            }
            flag = false;
            for(int m=0; m < matched.size(); m++) {
                if(directory.get(keySet.get(i)).equals(matched.get(m))) {
                    flag = true;
                }
            }
            if(flag == false) {
                adjustedDir.put(directory.get(keySet.get(i)), blocks);
                blocks = "";
            }
            if(flag == false){
                for(int k=(i+1); k < directory.size(); k++) {
                    if(directory.get(keySet.get(i)).equals(directory.get(keySet.get(k)))) {
                        block = (Integer) keySet.get(k);
                        matched.add(directory.get(keySet.get(i)));
                        for(int l=0; l < (Integer) allocList.get(keySet.get(k)); l++) {
                            blocks += block + " ";
                            block++;
                        }
                    }
                }
            }
            if(!blocks.equals("")) {
                blocks += adjustedDir.get(directory.get(keySet.get(i)));
                adjustedDir.replace(directory.get(keySet.get(i)), blocks);
            }
        }
    }

    public void createFileNumbers() {
        number = 1;
        fileNumbers = new HashMap();
        keySet = new ArrayList();
        for ( Object key : adjustedDir.keySet() ) {
            keySet.add(key);
        }
        for(int i=0; i < adjustedDir.size(); i++) {
            fileNumbers.put(keySet.get(i), number);
            number++;
        }
    }

    public void createDetailsArray() {
        detailsArray = new int[size];
        number = 1;
        Iterator keySetIterator = adjustedDir.keySet().iterator();
        while (keySetIterator.hasNext())
        {
            String key = (String) keySetIterator.next();
            String str = (String) adjustedDir.get(key);
            String[] tokens = str.split(" ");
            for(int i=0; i < tokens.length; i++) {
                detailsArray[Integer.parseInt(tokens[i])] = number;
            }
            number++;
        }
    }

    public void printDir() {
        keySet = new ArrayList();
        adjustedDir = new HashMap();
        matched = new ArrayList();
        System.out.println();
        System.out.println("============== Current Drive Contents =================");
        System.out.println();
        System.out.println("DIRECTORY:");
        createAdjustedDirectory();

        // Print adjusted directory
        number = 1;
        keySet = new ArrayList();
        for ( Object key : adjustedDir.keySet() ) {
            keySet.add(key);
        }
        for(int i=0; i < adjustedDir.size(); i++) {
            System.out.print(number + ". " + keySet.get(i) + ", Block(s) ");
            System.out.print(adjustedDir.get(keySet.get(i)));
            System.out.println();
            number++;
        }
        System.out.println();

        System.out.println("DETAILS:");
        createDetailsArray();

        // Print the details array
        for(int i=0; i < detailsArray.length; i++) {
            if((i % 10) == 0 && i != 0)
                System.out.println();
            if(detailsArray[i] == 0)
                System.out.print("* ");
            else
                System.out.print(detailsArray[i] + " ");
        }
        System.out.println();
        System.out.println();
    }

    public void printStats() {
        System.out.println();
        System.out.println("During this simulation,");
        System.out.println("Total head moves = " + totMoves);
        System.out.println("Total number of files not " +
            "allocated due to insufficient space = " + notAlloc);
    }

    public void read(String fileName) {
        numMoves = 0;
        createAdjustedDirectory();
        createFileNumbers();
        createDetailsArray();

        // calculate numMoves
        for(int i=0; i < detailsArray.length; i++) {

            if(detailsArray[i] == (Integer) fileNumbers.get(fileName)) {
                numMoves++;
            }
            while(detailsArray[i] == (Integer) fileNumbers.get(fileName)) {
                if(i == detailsArray.length-1) {
                    totMoves += numMoves;
                    System.out.println("File " + fileName + " was read successfully with " + numMoves + " head move(s)");
                    return;
                }
                i++;
            }
        }
        totMoves += numMoves;
        System.out.println("File " + fileName + " was read successfully with " + numMoves + " head move(s)");
    }

    public boolean contiguous(String fileName, int sizeOfFile) {
        temp = new int[size];
        count = 0;
        for(Object value: tempList.values()) {
            temp[count] = (Integer) value;
            count++;
        }
        sort(temp); // sort the holes smallest to largest
        // check for perfect fit
        for(int i=0; i < size; i++) {
            if(tempList.get(i) != null) {
                if((Integer) tempList.get(i) == sizeOfFile){
                    allocList.put(i, sizeOfFile);
                    directory.put(i, fileName);
                    System.out.println("File " + fileName + " was added successfully");
                    return true;
                }
            }
        }

        // check next largest hole 
        for(int i=0; i < temp.length; i++) {
            if(temp[i] != 0 && temp[i] > sizeOfFile) { 
                for(int j=0; j < size; j++) {
                    if(tempList.get(j) != null) {
                        if((Integer) tempList.get(j) == temp[i]) {
                            allocList.put(j, sizeOfFile); // add to allocList
                            directory.put(j, fileName); // add to directory
                            System.out.println("File " + fileName + " was added successfully");
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    public boolean indexed(String fileName, int sizeOfFile) {
        // find total number of holes
        totalHoles = 0;
        for(int i=0;i < size; i++) {
            if(tempList.get(i) != null) {
                totalHoles += (Integer) tempList.get(i);
            }
        }

        // Calculate FAT blocks
        sizeOfFile = calcFAT(sizeOfFile);
        // Able to allocate
        if(sizeOfFile <= totalHoles) {
            for (int i = 0; i < size; i++) {
                if (tempList.get(i) != null) {
                    if ((Integer) tempList.get(i) < sizeOfFile) {
                        allocList.put(i, tempList.get(i));
                        directory.put(i, fileName);
                        sizeOfFile = sizeOfFile - (Integer) tempList.get(i);
                    } else {
                        allocList.put(i, sizeOfFile);
                        directory.put(i, fileName);
                        System.out.println("File " + fileName + " was added successfully");
                        return true;
                    }
                }
            }
        }
        return false;
    }

    public int calcFAT(int sizeOfFile) {
        fatBlocks = 0;
        for(int i=0; i < sizeOfFile; i++) {
            if((i % 7) == 0) {
                fatBlocks++;
            }
        }
        return sizeOfFile + fatBlocks;
    }
}