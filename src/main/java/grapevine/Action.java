package grapevine;

public class Action {
    public String charName;
    public Date actDate;
    public ActionNode subAction;
    public boolean done;
    public Date lastModified;

    private ActionNode firstNode;
    private ActionNode lastNode;
    private int NodeCount;
}