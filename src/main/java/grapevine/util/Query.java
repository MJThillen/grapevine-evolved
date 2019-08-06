package grapevine.util;

import grapevine.model.QueryClause;
import grapevine.constants.QueryConstants.*;

import java.time.LocalDate;
import java.util.LinkedList;

public class Query {
    private String name;
    private QueryInventory inventory;
    private boolean matchAll;
    private String sortKey;
    private boolean sortDescend;
    private LocalDate lastModified;
    private LinkedList<QueryClause> clauses;

    public Query() {
        clauses = new LinkedList<>();
    }

    public Query(String name,
                 QueryInventory inventory,
                 boolean matchAll, String sortKey,
                 boolean sortDescend,
                 LocalDate lastModified,
                 LinkedList<QueryClause> clauses) {
        this.name = name;
        this.inventory = inventory;
        this.matchAll = matchAll;
        this.sortKey = sortKey;
        this.sortDescend = sortDescend;
        this.lastModified = lastModified;
        this.clauses = clauses;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public QueryInventory getInventory() {
        return inventory;
    }

    public void setInventory(QueryInventory inventory) {
        this.inventory = inventory;
    }

    public boolean isMatchAll() {
        return matchAll;
    }

    public void setMatchAll(boolean matchAll) {
        this.matchAll = matchAll;
    }

    public String getSortKey() {
        return sortKey;
    }

    public void setSortKey(String sortKey) {
        this.sortKey = sortKey;
    }

    public boolean isSortDescend() {
        return sortDescend;
    }

    public void setSortDescend(boolean sortDescend) {
        this.sortDescend = sortDescend;
    }

    public LocalDate getLastModified() {
        return lastModified;
    }

    public void setLastModified(LocalDate lastModified) {
        this.lastModified = lastModified;
    }

    public LinkedList<QueryClause> getClauses() {
        return clauses;
    }

    public void setClauses(LinkedList<QueryClause> clauses) {
        this.clauses = clauses;
    }

    public void addClause(QueryClause clause) {
        clauses.add(clause);
    }
}
