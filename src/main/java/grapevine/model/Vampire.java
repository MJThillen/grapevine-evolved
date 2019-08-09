package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Vampire extends Character {
    private int generation;
    private String title;
    private String coterie;
    private String path;
    private String sire;
    private String aura;
    private String auraBonus;

    private int blood;
    private int conscience;
    private int selfControl;
    private int courage;
    private int pathTraits;

    private TraitList<Trait> bonds;
    private List<Boon> boons;
    private TraitList<Trait> boonTraits;
    private TraitList<Trait> derangements;
    private TraitList<Trait> disciplines;
    private TraitList<Trait> rituals;
    private TraitList<Trait> statusList;
    private TraitList<Trait> miscellaneous;

    public Vampire() {
        super(Race.VAMPIRE);
    }

    public Vampire(Player player,
                   String status,
                   LocalDate startDate,
                   Experience experience,
                   String narrator,
                   boolean isNPC,
                   Race race,
                   int generation,
                   String title,
                   String coterie,
                   String path,
                   String sire,
                   String aura,
                   String auraBonus,
                   int blood,
                   int conscience,
                   int selfControl,
                   int courage,
                   int pathTraits,
                   TraitList<Trait> bonds,
                   List<Boon> boons,
                   TraitList<Trait> boonTraits,
                   TraitList<Trait> derangements,
                   TraitList<Trait> disciplines,
                   TraitList<Trait> rituals,
                   TraitList<Trait> statusList,
                   TraitList<Trait> miscellaneous) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.generation = generation;
        this.title = title;
        this.coterie = coterie;
        this.path = path;
        this.sire = sire;
        this.aura = aura;
        this.auraBonus = auraBonus;
        this.blood = blood;
        this.conscience = conscience;
        this.selfControl = selfControl;
        this.courage = courage;
        this.pathTraits = pathTraits;
        this.bonds = bonds;
        this.boons = boons;
        this.boonTraits = boonTraits;
        this.derangements = derangements;
        this.disciplines = disciplines;
        this.rituals = rituals;
        this.statusList = statusList;
        this.miscellaneous = miscellaneous;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the clan.
     * */
    public String getClan() {
        return this.getGroup();
    }

    /**
     * A wrapper for readability
     * @param clan the value of clan to store in the group variable
     */
    public void setClan(String clan) {
        this.setGroup(clan);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the sect.
     */
    public String getSect() {
        return this.getSubGroup();
    }

    /**
     * A wrapper for readability
     * @param sect the value of sect to store in the subGroup variable
     */
    public void setSect(String sect) {
        this.setSubGroup(sect);
    }

    /* Generic Getters and Setters */

    public int getGeneration() {
        return generation;
    }

    public void setGeneration(int generation) {
        this.generation = generation;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getCoterie() {
        return coterie;
    }

    public void setCoterie(String coterie) {
        this.coterie = coterie;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }

    public String getSire() {
        return sire;
    }

    public void setSire(String sire) {
        this.sire = sire;
    }

    public String getAura() {
        return aura;
    }

    public void setAura(String aura) {
        this.aura = aura;
    }

    public String getAuraBonus() {
        return auraBonus;
    }

    public void setAuraBonus(String auraBonus) {
        this.auraBonus = auraBonus;
    }

    public int getBlood() {
        return blood;
    }

    public void setBlood(int blood) {
        this.blood = blood;
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

    public int getPathTraits() {
        return pathTraits;
    }

    public void setPathTraits(int pathTraits) {
        this.pathTraits = pathTraits;
    }

    public TraitList<Trait> getBonds() {
        return bonds;
    }

    public void setBonds(TraitList<Trait> bonds) {
        this.bonds = bonds;
    }

    public List<Boon> getBoons() {
        return boons;
    }

    public void setBoons(List<Boon> boons) {
        this.boons = boons;
    }

    public TraitList<Trait> getBoonTraits() {
        return boonTraits;
    }

    public void setBoonTraits(TraitList<Trait> boonTraits) {
        this.boonTraits = boonTraits;
    }

    public TraitList<Trait> getDerangements() {
        return derangements;
    }

    public void setDerangements(TraitList<Trait> derangements) {
        this.derangements = derangements;
    }

    public TraitList<Trait> getDisciplines() {
        return disciplines;
    }

    public void setDisciplines(TraitList<Trait> disciplines) {
        this.disciplines = disciplines;
    }

    public TraitList<Trait> getRituals() {
        return rituals;
    }

    public void setRituals(TraitList<Trait> rituals) {
        this.rituals = rituals;
    }

    public TraitList<Trait> getStatusList() {
        return statusList;
    }

    public void setStatusList(TraitList<Trait> statusList) {
        this.statusList = statusList;
    }

    public TraitList<Trait> getMiscellaneous() {
        return miscellaneous;
    }

    public void setMiscellaneous(TraitList<Trait> miscellaneous) {
        this.miscellaneous = miscellaneous;
    }
}
