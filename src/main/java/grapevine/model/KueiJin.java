package grapevine.model;

public class KueiJin extends Character {
    private String direction;
    private String station;
    private String poArchetype;

    private int hun;
    private int po;
    private int yinChi;
    private int yangChi;
    private int demonChi;
    private int dharmaTraits;

    private TraitList<Trait> statusList;
    private TraitList<Trait> guanxi;
    private TraitList<Trait> disciplines;
    private TraitList<Trait> rites;

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the dharma.
     * */
    public String getDharma() {
        return this.getGroup();
    }

    /**
     * A wrapper for readability
     * @param dharma the value of dharma to store in the group variable
     */
    public void setDharma(String dharma) {
        this.setGroup(dharma);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the balance.
     */
    public String getBalance() {
        return this.getSubGroup();
    }

    /**
     * A wrapper for readability
     * @param balance the value of sect to store in the subGroup variable
     */
    public void setBalance(String balance) {
        this.setSubGroup(balance);
    }

    /*  Generic Getters and Setters */
}
