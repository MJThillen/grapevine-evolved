package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Fera extends Character {
    private String rank;
    private String breed;
    private String pack;
    private String totem;
    private String position;

    private int notoriety;
    private String honor;
    private String glory;
    private String wisdom;
    private TraitList<Trait> features;
    private TraitList<Trait> gifts;
    private TraitList<Trait> rites;

    public Fera() {
        super(Race.FERA);
    }

    public Fera(Player player,
                String status,
                LocalDate startDate,
                Experience experience,
                String narrator,
                boolean isNPC,
                Race race,
                String rank,
                String breed,
                String pack,
                String totem,
                String position,
                int notoriety,
                String honor,
                String glory,
                String wisdom,
                TraitList<Trait> features,
                TraitList<Trait> gifts,
                TraitList<Trait> rites) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.rank = rank;
        this.breed = breed;
        this.pack = pack;
        this.totem = totem;
        this.position = position;
        this.notoriety = notoriety;
        this.honor = honor;
        this.glory = glory;
        this.wisdom = wisdom;
        this.features = features;
        this.gifts = gifts;
        this.rites = rites;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the fera.
     * */
    public String getFera() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param fera The value of fera to store in the group variable.
     */
    public void setFera(String fera) {
        this.setGroup(fera);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the auspice.
     */
    public String getAuspice() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param auspice the value of auspice to store in the subgroup variable.
     */
    public void setAuspice(String auspice) {
        this.setSubGroup(auspice);
    }

    /* Generic Getters and Setters */

    public String getRank() {
        return rank;
    }

    public void setRank(String rank) {
        this.rank = rank;
    }

    public String getBreed() {
        return breed;
    }

    public void setBreed(String breed) {
        this.breed = breed;
    }

    public String getPack() {
        return pack;
    }

    public void setPack(String pack) {
        this.pack = pack;
    }

    public String getTotem() {
        return totem;
    }

    public void setTotem(String totem) {
        this.totem = totem;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    public int getNotoriety() {
        return notoriety;
    }

    public void setNotoriety(int notoriety) {
        this.notoriety = notoriety;
    }

    public String getHonor() {
        return honor;
    }

    public void setHonor(String honor) {
        this.honor = honor;
    }

    public String getGlory() {
        return glory;
    }

    public void setGlory(String glory) {
        this.glory = glory;
    }

    public String getWisdom() {
        return wisdom;
    }

    public void setWisdom(String wisdom) {
        this.wisdom = wisdom;
    }

    public TraitList<Trait> getFeatures() {
        return features;
    }

    public void setFeatures(TraitList<Trait> features) {
        this.features = features;
    }

    public TraitList<Trait> getGifts() {
        return gifts;
    }

    public void setGifts(TraitList<Trait> gifts) {
        this.gifts = gifts;
    }

    public TraitList<Trait> getRites() {
        return rites;
    }

    public void setRites(TraitList<Trait> rites) {
        this.rites = rites;
    }
}
