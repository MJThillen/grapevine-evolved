package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Mortal extends Character {
    private String title;
    private String regnant;
    private int blood;
    private int humanity;
    private int conscience;
    private int selfControl;
    private int courage;
    private int trueFaith;
    private TraitList<Trait> humanities;
    private TraitList<Trait> derangements;
    private TraitList<Trait> numina;

    public Mortal() {
        super(Race.MORTAL);
    }

    public Mortal(Player player,
                  String status,
                  LocalDate startDate,
                  Experience experience,
                  String narrator,
                  boolean isNPC,
                  Race race,
                  String title,
                  String regnant,
                  int blood,
                  int humanity,
                  int conscience,
                  int selfControl,
                  int courage,
                  int trueFaith,
                  TraitList<Trait> humanities,
                  TraitList<Trait> derangements,
                  TraitList<Trait> numina) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.title = title;
        this.regnant = regnant;
        this.blood = blood;
        this.humanity = humanity;
        this.conscience = conscience;
        this.selfControl = selfControl;
        this.courage = courage;
        this.trueFaith = trueFaith;
        this.humanities = humanities;
        this.derangements = derangements;
        this.numina = numina;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the association.
     * */
    public String getAssociation() {
        return this.getGroup();
    }

    /**
     * A wrapper for readability
     * @param association the value of association to store in the group variable
     */
    public void setAssociation(String association) {
        this.setGroup(association);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the motivation.
     */
    public String getMotivation() {
        return this.getSubGroup();
    }

    /**
     * A wrapper for readability
     * @param motivation the value of motivation to store in the subGroup variable
     */
    public void setMotivation(String motivation) {
        this.setSubGroup(motivation);
    }

    /* Generic Getters and Setters */

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getRegnant() {
        return regnant;
    }

    public void setRegnant(String regnant) {
        this.regnant = regnant;
    }

    public int getBlood() {
        return blood;
    }

    public void setBlood(int blood) {
        this.blood = blood;
    }

    public int getHumanity() {
        return humanity;
    }

    public void setHumanity(int humanity) {
        this.humanity = humanity;
    }

    public int getConscience() {
        return conscience;
    }

    public void setConscience(int conscience) {
        this.conscience = conscience;
    }

    public int getSelfControl() {
        return selfControl;
    }

    public void setSelfControl(int selfControl) {
        this.selfControl = selfControl;
    }

    public int getCourage() {
        return courage;
    }

    public void setCourage(int courage) {
        this.courage = courage;
    }

    public int getTrueFaith() {
        return trueFaith;
    }

    public void setTrueFaith(int trueFaith) {
        this.trueFaith = trueFaith;
    }

    public TraitList<Trait> getHumanities() {
        return humanities;
    }

    public void setHumanities(TraitList<Trait> humanities) {
        this.humanities = humanities;
    }

    public TraitList<Trait> getDerangements() {
        return derangements;
    }

    public void setDerangements(TraitList<Trait> derangements) {
        this.derangements = derangements;
    }

    public TraitList<Trait> getNumina() {
        return numina;
    }

    public void setNumina(TraitList<Trait> numina) {
        this.numina = numina;
    }
}
