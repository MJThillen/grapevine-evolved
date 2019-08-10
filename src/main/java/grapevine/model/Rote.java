package grapevine.model;

import java.time.LocalDate;

public class Rote {
    private String name;
    private int level;
    private String duration;
    private TraitList<Trait> spheres;
    private String description;
    private String grades;
    private String iconKey;
    private LocalDate lastModified;

    public Rote(String name,
                int level,
                String duration,
                TraitList<Trait> spheres,
                String description,
                String grades,
                String iconKey,
                LocalDate lastModified) {
        this.name = name;
        this.level = level;
        this.duration = duration;
        this.spheres = spheres;
        this.description = description;
        this.grades = grades;
        this.iconKey = iconKey;
        this.lastModified = lastModified;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getLevel() {
        return level;
    }

    public void setLevel(int level) {
        this.level = level;
    }

    public String getDuration() {
        return duration;
    }

    public void setDuration(String duration) {
        this.duration = duration;
    }

    public TraitList<Trait> getSpheres() {
        return spheres;
    }

    public void setSpheres(TraitList<Trait> spheres) {
        this.spheres = spheres;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getGrades() {
        return grades;
    }

    public void setGrades(String grades) {
        this.grades = grades;
    }

    public String getIconKey() {
        return iconKey;
    }

    public void setIconKey(String iconKey) {
        this.iconKey = iconKey;
    }

    public LocalDate getLastModified() {
        return lastModified;
    }

    public void setLastModified(LocalDate lastModified) {
        this.lastModified = lastModified;
    }
}
