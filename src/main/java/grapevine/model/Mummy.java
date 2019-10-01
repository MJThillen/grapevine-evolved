package grapevine.model;

public class Mummy extends Character {
    private int sekhem;
    private int balance;
    private int memory;
    private int joy;
    private int ba;
    private int ka;

    private TraitList<Trait> statusList;
    private TraitList<Trait> humanityList;
    private TraitList<Trait> hekauList;
    private TraitList<Trait> spellList;
    private TraitList<Trait> ritualList;

    public int getSekhem() {
        return sekhem;
    }

    public void setSekhem(int sekhem) {
        this.sekhem = sekhem;
    }

    public int getBalance() {
        return balance;
    }

    public void setBalance(int balance) {
        this.balance = balance;
    }

    public int getMemory() {
        return memory;
    }

    public void setMemory(int memory) {
        this.memory = memory;
    }

    public int getJoy() {
        return joy;
    }

    public void setJoy(int joy) {
        this.joy = joy;
    }

    public int getBa() {
        return ba;
    }

    public void setBa(int ba) {
        this.ba = ba;
    }

    public int getKa() {
        return ka;
    }

    public void setKa(int ka) {
        this.ka = ka;
    }

    public TraitList<Trait> getStatusList() {
        return statusList;
    }

    public void setStatusList(TraitList<Trait> statusList) {
        this.statusList = statusList;
    }

    public TraitList<Trait> getHumanityList() {
        return humanityList;
    }

    public void setHumanityList(TraitList<Trait> humanityList) {
        this.humanityList = humanityList;
    }

    public TraitList<Trait> getHekauList() {
        return hekauList;
    }

    public void setHekauList(TraitList<Trait> hekauList) {
        this.hekauList = hekauList;
    }

    public TraitList<Trait> getSpellList() {
        return spellList;
    }

    public void setSpellList(TraitList<Trait> spellList) {
        this.spellList = spellList;
    }

    public TraitList<Trait> getRitualList() {
        return ritualList;
    }

    public void setRitualList(TraitList<Trait> ritualList) {
        this.ritualList = ritualList;
    }
    /* Generic getters and setters */
}
