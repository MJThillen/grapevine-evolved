package grapevine.model;

import grapevine.constants.Race;

import java.time.LocalDate;

public class Wraith extends Character {
    private String ethnos;
    private String faction;
    private String legion;
    private String rank;
    private String passions;
    private String fetters;
    private String life;
    private String death;
    private String haunt;
    private String regret;
    private String shadowArchetype;
    private String shadowPlayer;

    private int pathos;
    private int corpus;
    private int angst;
    private int darkPassions;

    private TraitList<Trait> statusList;
    private TraitList<Trait> influences;
    private TraitList<Trait> arcanoiList;
    private TraitList<Trait> thorns;

    public Wraith() {
        super(Race.WRAITH);
    }

    public Wraith(Player player,
                  String status,
                  LocalDate startDate,
                  Experience experience,
                  String narrator,
                  boolean isNPC,
                  Race race,
                  String ethnos,
                  String faction,
                  String legion,
                  String rank,
                  String passions,
                  String fetters,
                  String life,
                  String death,
                  String haunt,
                  String regret,
                  String shadowArchetype,
                  String shadowPlayer,
                  int pathos,
                  int corpus,
                  int angst,
                  int darkPassions,
                  TraitList<Trait> statusList,
                  TraitList<Trait> influences,
                  TraitList<Trait> arcanoiList,
                  TraitList<Trait> thorns) {
        super(player, status, startDate, experience, narrator, isNPC, race);
        this.ethnos = ethnos;
        this.faction = faction;
        this.legion = legion;
        this.rank = rank;
        this.passions = passions;
        this.fetters = fetters;
        this.life = life;
        this.death = death;
        this.haunt = haunt;
        this.regret = regret;
        this.shadowArchetype = shadowArchetype;
        this.shadowPlayer = shadowPlayer;
        this.pathos = pathos;
        this.corpus = corpus;
        this.angst = angst;
        this.darkPassions = darkPassions;
        this.statusList = statusList;
        this.influences = influences;
        this.arcanoiList = arcanoiList;
        this.thorns = thorns;
    }

    /**
     * A wrapper method for readability.
     * @return The group value where we stored the ethnos.
     * */
    public String getEthnos() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param ethnos The value of ethnos to store in the group variable.
     */
    public void setEthnos(String ethnos) {
        this.setGroup(ethnos);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the guild.
     */
    public String getGuild() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param guild the value of guild to store in the subgroup variable.
     */
    public void setGuild(String guild) {
        this.setSubGroup(guild);
    }

    /* Generic Getters and Setters */

    public String getFaction() {
        return faction;
    }

    public void setFaction(String faction) {
        this.faction = faction;
    }

    public String getLegion() {
        return legion;
    }

    public void setLegion(String legion) {
        this.legion = legion;
    }

    public String getRank() {
        return rank;
    }

    public void setRank(String rank) {
        this.rank = rank;
    }

    public String getPassions() {
        return passions;
    }

    public void setPassions(String passions) {
        this.passions = passions;
    }

    public String getFetters() {
        return fetters;
    }

    public void setFetters(String fetters) {
        this.fetters = fetters;
    }

    public String getLife() {
        return life;
    }

    public void setLife(String life) {
        this.life = life;
    }

    public String getDeath() {
        return death;
    }

    public void setDeath(String death) {
        this.death = death;
    }

    public String getHaunt() {
        return haunt;
    }

    public void setHaunt(String haunt) {
        this.haunt = haunt;
    }

    public String getRegret() {
        return regret;
    }

    public void setRegret(String regret) {
        this.regret = regret;
    }

    public String getShadowArchetype() {
        return shadowArchetype;
    }

    public void setShadowArchetype(String shadowArchetype) {
        this.shadowArchetype = shadowArchetype;
    }

    public String getShadowPlayer() {
        return shadowPlayer;
    }

    public void setShadowPlayer(String shadowPlayer) {
        this.shadowPlayer = shadowPlayer;
    }

    public int getPathos() {
        return pathos;
    }

    public void setPathos(int pathos) {
        this.pathos = pathos;
    }

    public int getCorpus() {
        return corpus;
    }

    public void setCorpus(int corpus) {
        this.corpus = corpus;
    }

    public int getAngst() {
        return angst;
    }

    public void setAngst(int angst) {
        this.angst = angst;
    }

    public int getDarkPassions() {
        return darkPassions;
    }

    public void setDarkPassions(int darkPassions) {
        this.darkPassions = darkPassions;
    }

    public TraitList<Trait> getStatusList() {
        return statusList;
    }

    public void setStatusList(TraitList<Trait> statusList) {
        this.statusList = statusList;
    }

    @Override
    public TraitList<Trait> getInfluences() {
        return influences;
    }

    @Override
    public void setInfluences(TraitList<Trait> influences) {
        this.influences = influences;
    }

    public TraitList<Trait> getArcanoiList() {
        return arcanoiList;
    }

    public void setArcanoiList(TraitList<Trait> arcanoiList) {
        this.arcanoiList = arcanoiList;
    }

    public TraitList<Trait> getThorns() {
        return thorns;
    }

    public void setThorns(TraitList<Trait> thorns) {
        this.thorns = thorns;
    }
}
