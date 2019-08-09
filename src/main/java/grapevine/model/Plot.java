package grapevine.model;

import java.io.Serializable;
import java.time.LocalDate;
import java.util.LinkedList;
import java.util.List;

public class Plot implements Serializable {
    private String name;
    private LocalDate date;
    private String outline;
    private String development;
    private List<Plot> causes;
    private List<Plot> effects;
    private LocalDate startDate;
    private LocalDate endDate;
    private LocalDate devDate;
    private String narrator;


    public Plot() {

    }

    public Plot(String name,
                LocalDate startDate,
                LocalDate endDate,
                String outline,
                String narrator,
                LocalDate devDate,
                String development,
                LinkedList<Plot> causes,
                LinkedList<Plot> effects) {
        this.setName(name);
        this.setStartDate(startDate);
        this.setEndDate(endDate);
        this.setDevDate(devDate);
        this.setOutline(outline);
        this.setDevelopment(development);
        this.setNarrator(narrator);
        this.setCauses(causes);
        this.setEffects(effects);
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

    public String getOutline() {
        return outline;
    }

    public void setOutline(String outline) {
        this.outline = outline;
    }

    public String getDevelopment() {
        return development;
    }

    public void setDevelopment(String development) {
        this.development = development;
    }

    public List<Plot> getCauses() {
        return causes;
    }

    public void setCauses(List<Plot> causes) {
        this.causes = causes;
    }

    public List<Plot> getEffects() {
        return effects;
    }

    public void setEffects(List<Plot> effects) {
        this.effects = effects;
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
