package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Various extends Character {
    private String affinity;
    private String plane;
    private String brood;
    private TraitList<Trait> tempers;
    private TraitList<Trait> powers;

    public Various() {
        super(Race.VARIOUS);
        this.tempers = new TraitList<Trait>();
        this.powers = new TraitList<Trait>();
    }

    public Various(Player player,
                   String status,
                   LocalDate startDate,
                   Experience experience,
                   String narrator,
                   boolean isNPC,
                   Race race,
                   String affinity,
                   String plane,
                   String brood,
                   TraitList<Trait> tempers,
                   TraitList<Trait> powers) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.affinity = affinity;
        this.plane = plane;
        this.brood = brood;
        this.tempers = tempers;
        this.powers = powers;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the charClass.
     * */
    public String getCharClass() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the subClass.
     */
    public String getSubClass() {
        return this.getSubGroup();
    }

    /* Generic Getters and Setters */

    public String getAffinity() {
        return affinity;
    }

    public void setAffinity(String affinity) {
        this.affinity = affinity;
    }

    public String getPlane() {
        return plane;
    }

    public void setPlane(String plane) {
        this.plane = plane;
    }

    public String getBrood() {
        return brood;
    }

    public void setBrood(String brood) {
        this.brood = brood;
    }

    public TraitList<Trait> getTempers() {
        return tempers;
    }

    public void setTempers(TraitList<Trait> tempers) {
        this.tempers = tempers;
    }

    public TraitList<Trait> getPowers() {
        return powers;
    }

    public void setPowers(TraitList<Trait> powers) {
        this.powers = powers;
    }
}
