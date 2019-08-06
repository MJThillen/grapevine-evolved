package grapevine.model;

import grapevine.constants.RumorCategory;
import grapevine.util.Query;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.Serializable;
import java.time.LocalDate;
import java.util.ArrayList;

import static grapevine.constants.Constants.*;
import static grapevine.constants.QueryConstants.*;

@XmlRootElement
public class Rumor extends Event implements Serializable {

    @XmlElement
    private RumorCategory category;
    @XmlElement
    private Query query;

    @XmlElement
    private boolean done;
    @XmlElement
    private LocalDate lastModified;

    public Rumor() {
        super(EVENT_TYPE.rumor);
        lastModified = LocalDate.now();
        this.setEffects(new ArrayList<>());
    }

    public void initializeQuery(final String what,
                                final LocalDate when,
                                final RumorCategory category) {
        query = new Query();
        query.setInventory(QueryInventory.CHARACTERS);
        this.setName(what);
        this.setDate(when);
        this.setCategory(category);
    }

    public void addClauseToQuery(QueryKeys key, String name, int number, QueryCompare comp, boolean isNot) {
        query.addClause(new QueryClause(key.getValue(), comp, name, number, isNot));
    }

    //I'd like to move things like this to a Controller class, eventually.
    public void initializeMulti(final String what,
                                final LocalDate when,
                                final String newKey,
                                final String newMatch) {
        //ToDo: What in the actual heck does this function do? Check Context. I did my best to imitate.
        this.setName(what);
        this.setDate(when);
        this.setCategory(RumorCategory.INFLUENCE);
        this.setEvent(newKey);
        this.setSubEvent(newMatch);
        Rumor subRumor = new Rumor();
        subRumor.setName(newMatch);
        subRumor.setCategory(RumorCategory.INFLUENCE);
        subRumor.setEvent(newKey);
        subRumor.setSubEvent(newMatch);
        for (int i = 1; i < 10; i++) {
            this.addEffect(subRumor);
        }

    }

    public int outputID() {
        return OUTPUT_ID_CONSTANTS.rumor.getValue();
    }

    public String getValue(final QueryKeys key) {
        switch(key) {
            case TITLE:
                return this.getName();
            case DATE:
                return this.getDate().toString();
            case COUNT:
                return "" + getEffects().size();
            case RUMOR:
                return (getEffects().size() > 0 ? getEffects().get(0).toString() : "");
            case LEVEL:
                return (category.equals(RumorCategory.INFLUENCE) && getEffects().size() > 0 ?
                        "" + getEffects().get(0).getLevel() : "");
            case TYPE:
                if (category.equals(RumorCategory.INFLUENCE)) {
                    return (getEffects().size() > 0 ?
                            "Level " + getEffects().get(0).getLevel() + " " + this.getName() :
                            this.getName());
                } else {
                    return this.getName();
                }
            default:
                return "";
        }
    }

    public String iconKey() {
        switch(this.category) {
            case INFLUENCE:
                return "InfluenceRumor";
            case GROUP:
                return "GroupRumor";
            case SUBGROUP:
                return "SubgroupRumor";
            case RACE:
                return "RaceRumor";
            case PERSONAL:
                return "PersonalRumor";
            default:
                return "Rumor";
        }
    }

    public String shortDesc() {
        return getDate() + " " + getName();
    }

    //Getters and Setters

    public RumorCategory getCategory() {
        return category;
    }

    public void setCategory(RumorCategory category) {
        this.category = category;
    }

    public Query getQuery() {
        return query;
    }

    public void setQuery(Query query) {
        this.query = query;
    }

    public boolean isDone() {
        return done;
    }

    public void setDone(boolean done) {
        this.done = done;
    }

    public LocalDate getLastModified() {
        return lastModified;
    }

    public void setLastModified(LocalDate lastModified) {
        this.lastModified = lastModified;
    }

}
