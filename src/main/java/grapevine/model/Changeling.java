package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Changeling extends Character {
    private String seelieLegacy;
    private String unseelieLegacy;
    private String court;
    private String house;
    private String title;
    private String threshold;

    private int glamour;
    private int banality;

    private TraitList<Trait> statusList;
    private TraitList<Trait> arts;
    private TraitList<Trait> realms;
    private TraitList<Trait> oaths;

    public Changeling() {
        super(Race.CHANGELING);
    }

    public Changeling(Player player,
                      String status,
                      LocalDate startDate,
                      Experience experience,
                      String narrator,
                      boolean isNPC,
                      Race race,
                      String seelieLegacy,
                      String unseelieLegacy,
                      String court,
                      String house,
                      String title,
                      String threshold,
                      int glamour,
                      int banality,
                      TraitList<Trait> statusList,
                      TraitList<Trait> arts,
                      TraitList<Trait> realms,
                      TraitList<Trait> oaths) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.seelieLegacy = seelieLegacy;
        this.unseelieLegacy = unseelieLegacy;
        this.court = court;
        this.house = house;
        this.title = title;
        this.threshold = threshold;
        this.glamour = glamour;
        this.banality = banality;
        this.statusList = statusList;
        this.arts = arts;
        this.realms = realms;
        this.oaths = oaths;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the clan.
     * */
    public String getKith() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param kith The value of kith to store in the group variable.
     */
    public void setKith(String kith) {
        this.setGroup(kith);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the sect.
     */
    public String getSeeming() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param seeming the value of seeming to store in the subgroup variable.
     */
    public void setSeeming(String seeming) {
        this.setSubGroup(seeming);
    }

    /* Generic Getters and Setters */

    public String getSeelieLegacy() {
        return seelieLegacy;
    }

    public void setSeelieLegacy(String seelieLegacy) {
        this.seelieLegacy = seelieLegacy;
    }

    public String getUnseelieLegacy() {
        return unseelieLegacy;
    }

    public void setUnseelieLegacy(String unseelieLegacy) {
        this.unseelieLegacy = unseelieLegacy;
    }

    public String getCourt() {
        return court;
    }

    public void setCourt(String court) {
        this.court = court;
    }

    public String getHouse() {
        return house;
    }

    public void setHouse(String house) {
        this.house = house;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getThreshold() {
        return threshold;
    }

    public void setThreshold(String threshold) {
        this.threshold = threshold;
    }

    public int getGlamour() {
        return glamour;
    }

    public void setGlamour(int glamour) {
        this.glamour = glamour;
    }

    public int getBanality() {
        return banality;
    }

    public void setBanality(int banality) {
        this.banality = banality;
    }

    public TraitList<Trait> getStatusList() {
        return statusList;
    }

    public void setStatusList(TraitList<Trait> statusList) {
        this.statusList = statusList;
    }

    public TraitList<Trait> getArts() {
        return arts;
    }

    public void setArts(TraitList<Trait> arts) {
        this.arts = arts;
    }

    public TraitList<Trait> getRealms() {
        return realms;
    }

    public void setRealms(TraitList<Trait> realms) {
        this.realms = realms;
    }

    public TraitList<Trait> getOaths() {
        return oaths;
    }

    public void setOaths(TraitList<Trait> oaths) {
        this.oaths = oaths;
    }
}
