package grapevine.model;

import grapevine.constants.Constants;

import java.io.Serializable;
import java.time.LocalDate;

public class Action extends Event implements Serializable {
    private String characterName;
    private int unused;
    private int total;
    private int growth;
    private boolean done;

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
        super(EVENT_TYPE.action);
        this.setName(name);
        this.setDate(date);
        this.setEvent(action);
        this.setSubEvent(result);
        this.setCharacterName(characterName);
        this.setLevel(level);
        this.setUnused(unused);
        this.setTotal(total);
        this.setGrowth(growth);
        this.setDone(done);
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