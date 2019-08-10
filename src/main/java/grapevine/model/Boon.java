package grapevine.model;

import java.time.LocalDate;

public class Boon {
    private String charName;
    private boolean isOwed;
    private String boonType;
    private LocalDate boonDate;
    private String description;

    public Boon(String charName,
                boolean isOwed,
                String boonType,
                LocalDate boonDate,
                String description) {
        this.charName = charName;
        this.isOwed = isOwed;
        this.boonType = boonType;
        this.boonDate = boonDate;
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

    public String getBoonType() {
        return boonType;
    }

    public void setBoonType(String boonType) {
        this.boonType = boonType;
    }

    public LocalDate getBoonDate() {
        return boonDate;
    }

    public void setBoonDate(LocalDate boonDate) {
        this.boonDate = boonDate;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }
}
