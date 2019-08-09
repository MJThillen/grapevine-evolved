package grapevine.model;

import java.io.Serializable;
import java.time.LocalDate;

public class Action implements Serializable {
    private String name;
    private LocalDate date;
    private String action;
    private String result;
    private int level;
    private String characterName;
    private int unused;
    private int total;
    private int growth;

    //Does this need to be persisted?
    private boolean done;

    public Action() {

    }

    public Action(String name,
                  LocalDate date,
                  String action,
                  String result,
                  String characterName,
                  int level,
                  int unused,
                  int total,
                  int growth,
                  boolean done) {
        this.setName(name);
        this.setDate(date);
        this.setAction(action);
        this.setResult(result);
        this.setCharacterName(characterName);
        this.setLevel(level);
        this.setUnused(unused);
        this.setTotal(total);
        this.setGrowth(growth);
        this.setDone(done);
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public LocalDate getDate() {
        return date;
    }

    public void setDate(LocalDate date) {
        this.date = date;
    }

    public String getAction() {
        return action;
    }

    public void setAction(String action) {
        this.action = action;
    }

    public String getResult() {
        return result;
    }

    public void setResult(String result) {
        this.result = result;
    }

    public int getLevel() {
        return level;
    }

    public void setLevel(int level) {
        this.level = level;
    }

    public String getCharacterName() {
        return characterName;
    }

    public void setCharacterName(String characterName) {
        this.characterName = characterName;
    }

    public int getUnused() {
        return unused;
    }

    public void setUnused(int unused) {
        this.unused = unused;
    }

    public int getTotal() {
        return total;
    }

    public void setTotal(int total) {
        this.total = total;
    }

    public int getGrowth() {
        return growth;
    }

    public void setGrowth(int growth) {
        this.growth = growth;
    }

    public boolean isDone() {
        return done;
    }

    public void setDone(boolean done) {
        this.done = done;
    }
}