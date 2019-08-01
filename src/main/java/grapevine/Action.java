package grapevine;

public class Action {
    public String charName;
    public Date actDate;
    public Event subAction;
    public boolean done;
    public Date lastModified;

    private Event firstNode;
    private Event lastNode;
    private int NodeCount;
}