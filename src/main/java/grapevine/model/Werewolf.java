package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Werewolf extends Character {
    private String breed;
    private String rank;
    private String pack;
    private String totem;
    private String camp;
    private String position;
    private int notoriety;
    private int rage;
    private int gnosis;
    private int honor;
    private int glory;
    private int wisdom;
    private TraitList<Trait> features;
    private TraitList<Trait> gifts;
    private TraitList<Trait> rites;
    private TraitList<Trait> renownList; //formerly 3 separate lists. May need to resplit.

    public Werewolf() {
        super(Race.WEREWOLF);
    }

    public Werewolf(Player player,
                    String status,
                    LocalDate startDate,
                    Experience experience,
                    String narrator,
                    boolean isNPC,
                    Race race,
                    String breed,
                    String rank,
                    String pack,
                    String totem,
                    String camp,
                    String position,
                    int notoriety,
                    int rage,
                    int gnosis,
                    int honor,
                    int glory,
                    int wisdom,
                    TraitList<Trait> features,
                    TraitList<Trait> gifts,
                    TraitList<Trait> rites,
                    TraitList<Trait> renownList) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.breed = breed;
        this.rank = rank;
        this.pack = pack;
        this.totem = totem;
        this.camp = camp;
        this.position = position;
        this.notoriety = notoriety;
        this.rage = rage;
        this.gnosis = gnosis;
        this.honor = honor;
        this.glory = glory;
        this.wisdom = wisdom;
        this.features = features;
        this.gifts = gifts;
        this.rites = rites;
        this.renownList = renownList;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the tribe.
     * */
    public String getTribe() {
        return this.getGroup();
    }

    /**
     * A wrapper for readability
     * @param tribe the value of tribe to store in the group variable
     */
    public void setTribe(String tribe) {
        this.setGroup(tribe);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the auspice.
     */
    public String getAuspice() {
        return this.getSubGroup();
    }

    /**
     * A wrapper for readability
     * @param auspice the value of auspice to store in the subGroup variable
     */
    public void setAuspice(String auspice) {
        this.setSubGroup(auspice);
    }

    /* Generic Getters and Setters */

    public String getBreed() {
        return breed;
    }

    public void setBreed(String breed) {
        this.breed = breed;
    }

    public String getRank() {
        return rank;
    }

    public void setRank(String rank) {
        this.rank = rank;
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

    public String getCamp() {
        return camp;
    }

    public void setCamp(String camp) {
        this.camp = camp;
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

    public int getRage() {
        return rage;
    }

    public void setRage(int rage) {
        this.rage = rage;
    }

    public int getGnosis() {
        return gnosis;
    }

    public void setGnosis(int gnosis) {
        this.gnosis = gnosis;
    }

    public int getHonor() {
        return honor;
    }

    public void setHonor(int honor) {
        this.honor = honor;
    }

    public int getGlory() {
        return glory;
    }

    public void setGlory(int glory) {
        this.glory = glory;
    }

    public int getWisdom() {
        return wisdom;
    }

    public void setWisdom(int wisdom) {
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

    public TraitList<Trait> getRenownList() {
        return renownList;
    }

    public void setRenownList(TraitList<Trait> renownList) {
        this.renownList = renownList;
    }
}
