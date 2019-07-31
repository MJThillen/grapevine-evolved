package grapevine.constants;

public enum ExperienceChange {
    EARNED(0),
    DEDUCTED(1),
    SET_EARNED(2),
    SPENT(3),
    UNSPENT(4),
    SET_UNSPENT(5),
    COMMENT(6);

    int change;

    ExperienceChange(int change) {
        this.change = change;
    }

    public int getChange() {
        return change;
    }
}
