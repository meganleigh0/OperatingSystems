import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * User: Megan Griffin
 * Date: 4/3/22
 */
public class PagingSimulation {

    public static void main(String[] args) {

        String FILE_NAME = "pages.dat";

        List<String> inputLines = new ArrayList<String>();

        InputClass input = new InputClass(FILE_NAME);
        try {
            input.parseInput(inputLines);
        } catch (FileNotFoundException e) {
            System.err.println("Could not find File");
            return;
        }

        for (int i = 0; i < inputLines.size(); i+=2) {
            int frames = Integer.valueOf(inputLines.get(i));

            List<Integer> tries = new ArrayList<Integer>();
            for (String s: inputLines.get(i+1).split(" ")) {
                tries.add(Integer.valueOf(s));
            }

            runCase(frames, tries);
        }
    }

    private static void runCase(int frames, List<Integer> tries) {
        Pager lru = Pager.getPager(Pager.PAGER_TYPE.LRU, frames, tries);
        Pager fifo = Pager.getPager(Pager.PAGER_TYPE.FIFO, frames, tries);
        Pager lfu = Pager.getPager(Pager.PAGER_TYPE.LFU, frames, tries);
        Pager opt = Pager.getPager(Pager.PAGER_TYPE.OPTIMAL, frames, tries);

        lru.execute();
        fifo.execute();
        lfu.execute();
        opt.execute();

        System.out.println("LRU: ");
        lru.printTable();

        System.out.println("FIFO: ");
        fifo.printTable();

        System.out.println("LFU: ");
        lfu.printTable();

        System.out.println("OPTIMAL: ");
        opt.printTable();
        
        System.out.println("Scheme \t\t#Faults \t%Optimal ");
        System.out.println("LRU \t\tFaults: " + lru.getFaults() + "\tPercent optimal: " + getPercentOptimal(lru, opt));
        System.out.println("FIFO \t\tFaults: " + fifo.getFaults() + "\tPercent optimal: " + getPercentOptimal(fifo, opt));
        System.out.println("LFU \t\tFaults: " + lfu.getFaults() + "\tPercent optimal: " + getPercentOptimal(lfu, opt));
        System.out.println("Optimal \tFaults: " + opt.getFaults() + "\tPercent optimal: " + getPercentOptimal(opt, opt));
    }

    private static String getPercentOptimal(Pager target, Pager opt) {
        DecimalFormat decimalFormat = new DecimalFormat("%###.#");
        return decimalFormat.format((double) target.getFaults() / (double) opt.getFaults());
    }

    private static class InputClass {

        String fileName;

        InputClass(String fileName) {
            this.fileName = fileName;
        }

        void parseInput(List<String> inputLines) throws FileNotFoundException {

            InputStream stream = null;
            try {

                stream = getClass().getClassLoader().getResourceAsStream(fileName);
                if (stream == null) {
                    System.err.println("File Not Found: " + fileName);
                    return;
                }

                StreamTokenizer tokenizer = new StreamTokenizer(new BufferedReader(new InputStreamReader(stream)));

                parseInput(inputLines, tokenizer);

            } catch (IOException e) {
                System.err.println("IOException while trying to read file.");
                e.printStackTrace();
            } finally {
                if (stream != null) {
                    try {
                        stream.close();
                    } catch (IOException e) {
                        System.err.println("Stream could not close");
                    }
                }
            }
        }

        void parseInput(List<String> inputLines, StreamTokenizer tokenizer) throws IOException {
            tokenizer.nextToken();
            while (tokenizer.ttype != StreamTokenizer.TT_EOF) {
                if (tokenizer.ttype == StreamTokenizer.TT_NUMBER) {
                    inputLines.add(String.valueOf((int) tokenizer.nval));
                    tokenizer.nextToken();

                    StringBuilder tryLine = new StringBuilder("");
                    while (tokenizer.ttype != StreamTokenizer.TT_EOF) {
                        if (tokenizer.nval > 0) {
                            tryLine.append((int) tokenizer.nval).append(" ");
                            tokenizer.nextToken();
                        } else {
                            tokenizer.nextToken();
                            break;
                        }
                    }
                    inputLines.add(tryLine.toString());
                }
            }
            tokenizer.nextToken();
        }
    }
}
