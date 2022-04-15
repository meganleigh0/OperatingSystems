/**
 * User: Megan Griffin
 * Date: 4/3/22
 */
public class Page {

    private int id;
    private int lastUsed = 0;

    private int accesses = 1;
    private int totalAccesses = 1;

    public Page(int id) {
        this.id = id;
    }

    public Page(int pageId, int round) {
        this.id = pageId;
        this.lastUsed = round;
    }

    public void incrementAccesses() {
        accesses++;
    }

    public void incrementOtherAccesses() {
        totalAccesses++;
    }

    public double getFrequency() {
        return accesses / totalAccesses;
    }

    public int getId() {
        return id;
    }

    public int getLastUsed() {
        return lastUsed;
    }

    public void setLastUsed(int lastUsed) {
        this.lastUsed = lastUsed;
    }
}