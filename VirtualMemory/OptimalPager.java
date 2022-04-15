import java.util.List;

/**
 * User: Megan Griffin
 * Date: 4/3/22
 */
public class OptimalPager extends Pager {

    public OptimalPager(int frameCount, Integer... tries) {
        super(frameCount, tries);
    }

    public OptimalPager(int frameCount, List<Integer> tries) {
        super(frameCount, tries);
    }

    @Override
    public void execute() {
        while (tries.size() > 0) {
            boolean fault = isPageFault(tries.get(0));
            if (fault) {
                handleFault(tries.get(0));
            } //else, do nothing
            takeStateSnapshot(tries.get(0), fault);
            tries.remove(0);
        }
    }

    private void handleFault(int pageId) {
        if (state.size() < frameCount) {
            state.add(new Page(pageId));
        } else {
            state.set(calculateReplacement(), new Page(pageId));
        }
    }

    private int calculateReplacement() {
        Page optimalPage = null;
        int nextUseOfOptimalPage = 0;
        for (Page p: state) {
            int nextUse = nextUseForPage(p.getId());

            if (nextUse < 0) {  //if the next use returned negative, the page is never used again, and so can be
                return state.indexOf(p); //replaced immediately
            } else if (nextUse > nextUseOfOptimalPage) {
                optimalPage = p;
                nextUseOfOptimalPage = nextUse;
            }
        }
        return optimalPage != null ? state.indexOf(optimalPage) : 0;
    }

    private int nextUseForPage(int pageId) {
        for (int i = 0; i < tries.size(); i++) {
            if (tries.get(i) == pageId) {
                return i;
            }
        }
        //if this page was not used again, it can be safely replaced without checking any further pages
        return -1;
    }
}