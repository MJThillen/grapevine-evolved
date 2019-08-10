package grapevine.model;

public class Mummy extends Character {
    private int Sekhem;
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

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the amenti.
     * */
    public String getAmenti() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param amenti The value of amenti to store in the group variable.
     */
    public void setAmenti(String amenti) {
        this.setGroup(amenti);
    }

    /**
     * A wrapper method for readability.
     * @return The subgroup value where we stored the inheritance.
     * */
    public String getInheritance() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param inheritance The value of inheritance to store in the subgroup variable.
     */
    public void setInheritance(String inheritance) {
        this.setSubGroup(inheritance);
    }


    /* Generic getters and setters */
}
