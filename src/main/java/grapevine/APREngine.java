package grapevine;

import java.util.Collections;
import java.util.LinkedList;

public class APREngine {
    private LinkedList<Event> actions;
    private LinkedList<Event> plots;
    private LinkedList<Event> rumors;
    private TraitList<Trait> traits;

    public APREngine() {
        actions = new LinkedList<>();
        plots = new LinkedList<>();
        rumors = new LinkedList<>();
        traits = new TraitList<>();
    }
}
