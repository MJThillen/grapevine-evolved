package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Hunter extends Character {
    private String handle;

    private int conviction;
    private int mercy;
    private int vision;
    private int zeal;

    private TraitList<Trait> derangements;
    private TraitList<Trait> edges;

    public Hunter() {
        super(Race.HUNTER);
    }

    public Hunter(Player player,
                  String status,
                  LocalDate startDate,
                  Experience experience,
                  String narrator,
                  boolean isNPC,
                  Race race,
                  String handle,
                  int conviction,
                  int mercy,
                  int vision,
                  int zeal,
                  TraitList<Trait> derangements,
                  TraitList<Trait> edges) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.handle = handle;
        this.conviction = conviction;
        this.mercy = mercy;
        this.vision = vision;
        this.zeal = zeal;
        this.derangements = derangements;
        this.edges = edges;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the creed.
     * */
    public String getCreed() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param creed The value of creed to store in the group variable.
     */
    public void setCreed(String creed) {
        this.setGroup(creed);
    }

    /**
     * A wrapper method for readability.
     * @return The subgroup value where we stored the camp.
     * */
    public String getCamp() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param camp The value of camp to store in the subgroup variable.
     */
    public void setCamp(String camp) {
        this.setSubGroup(camp);
    }

    /* Generic getters and setters */

    public String getHandle() {
        return handle;
    }

    public void setHandle(String handle) {
        this.handle = handle;
    }

    public int getConviction() {
        return conviction;
    }

    public void setConviction(int conviction) {
        this.conviction = conviction;
    }

    public int getMercy() {
        return mercy;
    }

    public void setMercy(int mercy) {
        this.mercy = mercy;
    }

    public int getVision() {
        return vision;
    }

    public void setVision(int vision) {
        this.vision = vision;
    }

    public int getZeal() {
        return zeal;
    }

    public void setZeal(int zeal) {
        this.zeal = zeal;
    }

    public TraitList<Trait> getDerangements() {
        return derangements;
    }

    public void setDerangements(TraitList<Trait> derangements) {
        this.derangements = derangements;
    }

    public TraitList<Trait> getEdges() {
        return edges;
    }

    public void setEdges(TraitList<Trait> edges) {
        this.edges = edges;
    }
}
