import java.util.Scanner;
import java.util.Arrays;
//Stats class
class Stats
{
    public static void main (String[] args)
    {
        Scanner sc = new Scanner(System.in);
        System.out.println ("Enter a single line of numbers: ");
        String line = sc.nextLine();
        String arr[] = line.split(" ");
        long nums[] = new long [arr.length];

        for(int i=0; i<arr.length; i++)
        {
            nums[i] = Long.parseLong(arr[i]);
        }

        AvgArray t0 = new AvgArray(nums);
        MinArray t1 = new MinArray(nums);
        MaxArray t2 = new MaxArray(nums);
        MedArray t3 = new MedArray(nums);
        SDArray t4 = new SDArray(nums);
        try{
            t0.join();
            t1.join();
            t2.join();
            t3.join();
            t4.join();
        }catch(Exception e){
            System.out.println (e);
        }

        System.out.println ("The average value is " + AvgArray.avg);
        System.out.println ("The minimum value is " + MinArray.min);
        System.out.println ("The maximum value is " + MaxArray.max);
        System.out.println ("The median value is " + MedArray.med);
        System.out.println ("The standard deviation value is " + SDArray.sd);
    }
}
//Average Array Thread
class AvgArray extends Thread
{
    private long arr[];
    public static long avg = 0;

    public AvgArray(long a[])
    {
        arr = new long[a.length];
        System.arraycopy(a, 0, arr, 0, a.length);
        start();
    }

    public void run()
    {
        long sum = 0;
        for(int i=0; i<arr.length; i++)
        {
            sum = sum + arr[i];
        }
        avg = sum / arr.length;
    }
}
//Minimum Array Thread
class MinArray extends Thread
{
    private long arr[];
    public static long min = 0;

    public MinArray(long a[])
    {
        arr = new long[a.length];
        System.arraycopy(a, 0, arr, 0, a.length);
        start();
    }

    public void run()
    {
        min = arr[0];
        for(int i=1; i<arr.length; i++)
        {
            if(min > arr[i])
                min = arr[i];
        }
    }
}
//Maximum Array Thread
class MaxArray extends Thread
{
    private long arr[];
    public static long max = 0;

    public MaxArray(long a[])
    {
        arr = new long[a.length];
        System.arraycopy(a, 0, arr, 0, a.length);
        start();
    }

    public void run()
    {
        max = arr[0];
        for(int i=1; i<arr.length; i++)
        {
            if(max < arr[i])
                max = arr[i];
        }
    }
}
//Median Array Thread
class MedArray extends Thread
{
    private long arr[];
    public static long med = 0;
    public MedArray(long a[])
    {
        Arrays.sort(a);
        int mid = a.length/2;
        if (a.length%2 == 1) {
            med = a[mid];
        } else {
            med = (a[mid-1] + a[mid])/2;
        }
    }
}
//Standard Deviation Array Thread
class SDArray extends Thread
{
    private long arr[];
    public static double sd = 0;

    public SDArray(long a[])
    {
        arr = new long[a.length];
        System.arraycopy(a, 0, arr, 0, a.length);
        start();
    }

    public void run()
    {
        long sum =0, avg = 0, length = arr.length;
        for(int i=0; i<arr.length; i++)
        {
            sum = sum + arr[i];
        }
        avg = sum / arr.length;
        for(double i: arr)
        {
            sd += Math.pow(i - avg, 2);
        }
        sd = Math.sqrt(sd/length);    
    }
}