import java.util.List;

/**
 * User: Megan Griffin
 * Date: 4/3/22
 */
public class LeastFrequentlyUsed extends Pager {

    public LeastFrequentlyUsed(int frameCount, Integer... tries) {
        super(frameCount, tries);
    }

    public LeastFrequentlyUsed(int frameCount, List<Integer> tries) {
        super(frameCount, tries);
    }

    @Override
    public void execute() {
        for (Integer pageId : tries) {
            boolean fault = isPageFault(pageId);
            if (fault) {
                handleFault(pageId);
            } else {
                for (Page p : state) {
                    p.incrementOtherAccesses();
                    if (p.getId() == pageId) {
                        p.incrementAccesses();
                    }
                }
            }
            takeStateSnapshot(pageId, fault);
        }
    }

    private void handleFault(int pageId) {
        if (state.size() < frameCount) {
            state.add(new Page(pageId));
        } else {

            double min = Double.MAX_VALUE;
            Page remove = null;
            for (Page p: state) {
                if (p.getFrequency() < min) {
                    min = p.getFrequency();
                    remove = p;
                }
                p.incrementOtherAccesses();
            }

            state.set(state.indexOf(remove), new Page(pageId));
        }
    }
}