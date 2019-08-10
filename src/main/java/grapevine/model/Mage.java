package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Mage extends Character {
    private String essence;
    private String faction;
    private String cabal;

    private int arete;
    private int quintessence;
    private int paradox;
    private TraitList<Trait> resonances;
    private TraitList<Trait> reputations;
    private TraitList<Trait> spheres;
    private TraitList<Trait> foci;
    private TraitList<Trait> rotes;

    public Mage() {
        super(Race.MAGE);
    }

    public Mage(Player player,
                String status,
                LocalDate startDate,
                Experience experience,
                String narrator,
                boolean isNPC,
                Race race,
                String essence,
                String faction,
                String cabal,
                int arete,
                int quintessence,
                int paradox,
                TraitList<Trait> resonances,
                TraitList<Trait> reputations,
                TraitList<Trait> spheres,
                TraitList<Trait> foci,
                TraitList<Trait> rotes) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.essence = essence;
        this.faction = faction;
        this.cabal = cabal;
        this.arete = arete;
        this.quintessence = quintessence;
        this.paradox = paradox;
        this.resonances = resonances;
        this.reputations = reputations;
        this.spheres = spheres;
        this.foci = foci;
        this.rotes = rotes;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the tradition.
     * */
    public String getTradition() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param tradition The value of tradition to store in the group variable.
     */
    public void setTradition(String tradition) {
        this.setGroup(tradition);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the rank.
     */
    public String getRank() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param rank the value of rank to store in the subgroup variable.
     */
    public void setRank(String rank) {
        this.setSubGroup(rank);
    }

    /* Generic Getters and Setters */

    public String getEssence() {
        return essence;
    }

    public void setEssence(String essence) {
        this.essence = essence;
    }

    public String getFaction() {
        return faction;
    }

    public void setFaction(String faction) {
        this.faction = faction;
    }

    public String getCabal() {
        return cabal;
    }

    public void setCabal(String cabal) {
        this.cabal = cabal;
    }

    public int getArete() {
        return arete;
    }

    public void setArete(int arete) {
        this.arete = arete;
    }

    public int getQuintessence() {
        return quintessence;
    }

    public void setQuintessence(int quintessence) {
        this.quintessence = quintessence;
    }

    public int getParadox() {
        return paradox;
    }

    public void setParadox(int paradox) {
        this.paradox = paradox;
    }

    public TraitList<Trait> getResonances() {
        return resonances;
    }

    public void setResonances(TraitList<Trait> resonances) {
        this.resonances = resonances;
    }

    public TraitList<Trait> getReputations() {
        return reputations;
    }

    public void setReputations(TraitList<Trait> reputations) {
        this.reputations = reputations;
    }

    public TraitList<Trait> getSpheres() {
        return spheres;
    }

    public void setSpheres(TraitList<Trait> spheres) {
        this.spheres = spheres;
    }

    public TraitList<Trait> getFoci() {
        return foci;
    }

    public void setFoci(TraitList<Trait> foci) {
        this.foci = foci;
    }

    public TraitList<Trait> getRotes() {
        return rotes;
    }

    public void setRotes(TraitList<Trait> rotes) {
        this.rotes = rotes;
    }
}
