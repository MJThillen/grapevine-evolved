package grapevine.model;

import java.time.LocalDate;

public class Boon {
    private String charName;
    private boolean isOwed;
    private String type;
    private LocalDate date;
    private String description;

    public Boon(String charName,
                boolean isOwed,
                String type,
                LocalDate date,
                String description) {
        this.charName = charName;
        this.isOwed = isOwed;
        this.type = type;
        this.date = date;
        this.description = description;
    }

    public String getCharName() {
        return charName;
    }

    public void setCharName(String charName) {
        this.charName = charName;
    }

    public boolean isOwed() {
        return isOwed;
    }

    public void setOwed(boolean owed) {
        isOwed = owed;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public LocalDate getDate() {
        return date;
    }

    public void setDate(LocalDate date) {
        this.date = date;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }
}
