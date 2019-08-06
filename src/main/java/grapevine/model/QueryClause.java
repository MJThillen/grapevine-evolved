package grapevine.model;

import grapevine.constants.QueryConstants.*;

public class QueryClause {
    private String key;
    private QueryCompare compareType;
    private String find;
    private double number;
    private boolean isNot;

    public QueryClause(String key, QueryCompare compareType, String find, double number, boolean isNot) {
        this.key = key;
        this.compareType = compareType;
        this.find = find;
        this.number = number;
        this.isNot = isNot;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public QueryCompare getCompareType() {
        return compareType;
    }

    public void setCompareType(QueryCompare compareType) {
        this.compareType = compareType;
    }

    public String getFind() {
        return find;
    }

    public void setFind(String find) {
        this.find = find;
    }

    public double getNumber() {
        return number;
    }

    public void setNumber(double number) {
        this.number = number;
    }

    public boolean isNot() {
        return isNot;
    }

    public void setNot(boolean not) {
        isNot = not;
    }
}
