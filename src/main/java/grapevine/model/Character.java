package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

/**
 * This is the parent class for all character types.  Shared values are declared and held here for ease of debugging.
 * Generic instances of Character are a terrible idea, don't do it ok?
 */
public class Character {
    //OOC Notes
    private int id;
    private Player player;
    private Game game;
    private String status;
    private LocalDate startDate;
    private Experience experience;
    private String narrator;
    private boolean isNPC;
    private LocalDate lastModified;

    //IC Information
    private Race race;
    private String name;
    private String nature;
    private String demeanor;
    private String group;
    private String subGroup;
    private String biography;
    private String notes;

    private int willpower;
    private int physicalMaxTraits;
    private int socialMaxTraits;
    private int mentalMaxTraits;

    private TraitList<Trait> physicalTraits;
    private TraitList<Trait> socialTraits;
    private TraitList<Trait> mentalTraits;

    private TraitList<Trait> physicalNegTraits;
    private TraitList<Trait> socialNegTraits;
    private TraitList<Trait> mentalNegTraits;

    private TraitList<Trait> abilities;
    private TraitList<Trait> influences;
    private TraitList<Trait> backgrounds;
    private TraitList<Trait> health;

    private TraitList<Trait> merits;
    private TraitList<Trait> flaws;

    private TraitList<Trait> equipment;
    private TraitList<Trait> hangouts;

    protected Character() {
    }

    protected Character(Race race) {
        this.race = race;
    }

    protected Character(Player player,
                     String status,
                     LocalDate startDate,
                     Experience experience,
                     String narrator,
                     boolean isNPC,
                     Race race) {
        this.player = player;
        this.status = status;
        this.startDate = startDate;
        this.experience = experience;
        this.narrator = narrator;
        this.isNPC = isNPC;
        this.race = race;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public Player getPlayer() {
        return player;
    }

    public void setPlayer(Player player) {
        this.player = player;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public LocalDate getStartDate() {
        return startDate;
    }

    public void setStartDate(LocalDate startDate) {
        this.startDate = startDate;
    }

    public Experience getExperience() {
        return experience;
    }

    public void setExperience(Experience experience) {
        this.experience = experience;
    }

    public String getNarrator() {
        return narrator;
    }

    public void setNarrator(String narrator) {
        this.narrator = narrator;
    }

    public boolean isNPC() {
        return isNPC;
    }

    public void setNPC(boolean NPC) {
        isNPC = NPC;
    }

    public LocalDate getLastModified() {
        return lastModified;
    }

    public void setLastModified(LocalDate lastModified) {
        this.lastModified = lastModified;
    }

    public Race getRace() {
        return race;
    }

    public void setRace(Race race) {
        this.race = race;
    }

    public String getGroup() {
        return group;
    }

    public void setGroup(String group) {
        this.group = group;
    }

    public String getSubGroup() {
        return subGroup;
    }

    public void setSubGroup(String subGroup) {
        this.subGroup = subGroup;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getNature() {
        return nature;
    }

    public void setNature(String nature) {
        this.nature = nature;
    }

    public String getDemeanor() {
        return demeanor;
    }

    public void setDemeanor(String demeanor) {
        this.demeanor = demeanor;
    }

    public String getBiography() {
        return biography;
    }

    public void setBiography(String biography) {
        this.biography = biography;
    }

    public String getNotes() {
        return notes;
    }

    public void setNotes(String notes) {
        this.notes = notes;
    }

    public int getWillpower() {
        return willpower;
    }

    public void setWillpower(int willpower) {
        this.willpower = willpower;
    }

    public int getPhysicalMaxTraits() {
        return physicalMaxTraits;
    }

    public void setPhysicalMaxTraits(int physicalMaxTraits) {
        this.physicalMaxTraits = physicalMaxTraits;
    }

    public int getSocialMaxTraits() {
        return socialMaxTraits;
    }

    public void setSocialMaxTraits(int socialMaxTraits) {
        this.socialMaxTraits = socialMaxTraits;
    }

    public int getMentalMaxTraits() {
        return mentalMaxTraits;
    }

    public void setMentalMaxTraits(int mentalMaxTraits) {
        this.mentalMaxTraits = mentalMaxTraits;
    }

    public TraitList<Trait> getPhysicalTraits() {
        return physicalTraits;
    }

    public void setPhysicalTraits(TraitList<Trait> physicalTraits) {
        this.physicalTraits = physicalTraits;
    }

    public TraitList<Trait> getSocialTraits() {
        return socialTraits;
    }

    public void setSocialTraits(TraitList<Trait> socialTraits) {
        this.socialTraits = socialTraits;
    }

    public TraitList<Trait> getMentalTraits() {
        return mentalTraits;
    }

    public void setMentalTraits(TraitList<Trait> mentalTraits) {
        this.mentalTraits = mentalTraits;
    }

    public TraitList<Trait> getPhysicalNegTraits() {
        return physicalNegTraits;
    }

    public void setPhysicalNegTraits(TraitList<Trait> physicalNegTraits) {
        this.physicalNegTraits = physicalNegTraits;
    }

    public TraitList<Trait> getSocialNegTraits() {
        return socialNegTraits;
    }

    public void setSocialNegTraits(TraitList<Trait> socialNegTraits) {
        this.socialNegTraits = socialNegTraits;
    }

    public TraitList<Trait> getMentalNegTraits() {
        return mentalNegTraits;
    }

    public void setMentalNegTraits(TraitList<Trait> mentalNegTraits) {
        this.mentalNegTraits = mentalNegTraits;
    }

    public TraitList<Trait> getAbilities() {
        return abilities;
    }

    public void setAbilities(TraitList<Trait> abilities) {
        this.abilities = abilities;
    }

    public TraitList<Trait> getInfluences() {
        return influences;
    }

    public void setInfluences(TraitList<Trait> influences) {
        this.influences = influences;
    }

    public TraitList<Trait> getBackgrounds() {
        return backgrounds;
    }

    public void setBackgrounds(TraitList<Trait> backgrounds) {
        this.backgrounds = backgrounds;
    }

    public TraitList<Trait> getHealth() {
        return health;
    }

    public void setHealth(TraitList<Trait> health) {
        this.health = health;
    }

    public TraitList<Trait> getMerits() {
        return merits;
    }

    public void setMerits(TraitList<Trait> merits) {
        this.merits = merits;
    }

    public TraitList<Trait> getFlaws() {
        return flaws;
    }

    public void setFlaws(TraitList<Trait> flaws) {
        this.flaws = flaws;
    }

    public TraitList<Trait> getEquipment() {
        return equipment;
    }

    public void setEquipment(TraitList<Trait> equipment) {
        this.equipment = equipment;
    }

    public TraitList<Trait> getHangouts() {
        return hangouts;
    }

    public void setHangouts(TraitList<Trait> hangouts) {
        this.hangouts = hangouts;
    }
}
