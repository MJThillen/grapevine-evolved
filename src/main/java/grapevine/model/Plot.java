package grapevine.model;

import java.io.Serializable;
import java.time.LocalDate;
import java.util.LinkedList;

public class Plot extends Event implements Serializable {
    private LocalDate startDate;
    private LocalDate endDate;
    private LocalDate devDate;
    private String narrator;


    public Plot() {
        super(EVENT_TYPE.plot);
    }

    public Plot(String name,
                LocalDate startDate,
                LocalDate endDate,
                String outline,
                String narrator,
                LocalDate devDate,
                String development,
                LinkedList<Event> causes,
                LinkedList<Event> effects) {
        super(EVENT_TYPE.plot);
        this.setName(name);
        this.setStartDate(startDate);
        this.setEndDate(endDate);
        this.setDevDate(devDate);
        this.setEvent(outline);
        this.setSubEvent(development);
        this.setNarrator(narrator);
        this.setCauses(causes);
        this.setEffects(effects);
    }

    public LocalDate getStartDate() {
        return startDate;
    }

    public void setStartDate(LocalDate startDate) {
        this.startDate = startDate;
    }

    public LocalDate getEndDate() {
        return endDate;
    }

    public void setEndDate(LocalDate endDate) {
        this.endDate = endDate;
    }

    public String getNarrator() {
        return narrator;
    }

    public void setNarrator(String narrator) {
        this.narrator = narrator;
    }

    public LocalDate getDevDate() {
        return devDate;
    }

    public void setDevDate(LocalDate devDate) {
        this.devDate = devDate;
    }

}
