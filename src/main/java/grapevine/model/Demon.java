package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Demon extends Character {
    private int torment;
    private int faith;
    private int conscience;
    private int conviction;
    private int courage;

    private TraitList<Trait> lore;
    private TraitList<Trait> visage;

    public Demon() {
        super(Race.DEMON);
    }

    public Demon(Player player,
                 String status,
                 LocalDate startDate,
                 Experience experience,
                 String narrator,
                 boolean isNPC,
                 Race race,
                 int torment,
                 int faith,
                 int conscience,
                 int conviction,
                 int courage,
                 TraitList<Trait> lore,
                 TraitList<Trait> visage) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.torment = torment;
        this.faith = faith;
        this.conscience = conscience;
        this.conviction = conviction;
        this.courage = courage;
        this.lore = lore;
        this.visage = visage;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the house.
     * */
    public String getHouse() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param house The value of house to store in the group variable.
     */
    public void setHouse(String house) {
        this.setGroup(house);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the faction.
     */
    public String getFaction() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param faction the value of rank to store in the subgroup variable.
     */
    public void setFaction(String faction) {
        this.setSubGroup(faction);
    }

    /* Generic getters and setters. */

    public int getTorment() {
        return torment;
    }

    public void setTorment(int torment) {
        this.torment = torment;
    }

    public int getFaith() {
        return faith;
    }

    public void setFaith(int faith) {
        this.faith = faith;
    }

    public int getConscience() {
        return conscience;
    }

    public void setConscience(int conscience) {
        this.conscience = conscience;
    }

    public int getConviction() {
        return conviction;
    }

    public void setConviction(int conviction) {
        this.conviction = conviction;
    }

    public int getCourage() {
        return courage;
    }

    public void setCourage(int courage) {
        this.courage = courage;
    }

    public TraitList<Trait> getLore() {
        return lore;
    }

    public void setLore(TraitList<Trait> lore) {
        this.lore = lore;
    }

    public TraitList<Trait> getVisage() {
        return visage;
    }

    public void setVisage(TraitList<Trait> visage) {
        this.visage = visage;
    }
}
