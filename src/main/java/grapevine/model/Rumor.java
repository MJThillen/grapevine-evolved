package grapevine.model;

import grapevine.constants.RumorCategory;
import grapevine.util.Query;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.Serializable;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import static grapevine.constants.Constants.*;
import static grapevine.constants.QueryConstants.*;

@XmlRootElement
public class Rumor implements Serializable {
    private String name;
    private LocalDate date;
    private String event;
    private String subEvent;
    private int level;
    private List<Rumor> subRumors;
    private RumorCategory category;
    private Query query;

    //do these need to exist?
    private boolean done;
    private LocalDate lastModified;

    public Rumor() {
        lastModified = LocalDate.now();
        this.setSubRumors(new ArrayList<>());
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
    public void influenceRumorSetup(final String what,
                                    final LocalDate when,
                                    final String newKey,
                                    final String newMatch) {
        //Influence Rumors have 10 levels, one for each possible level in the influence that a character can have.
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
            this.addSubRumor(subRumor);
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
                return "" + getSubRumors().size();
            case RUMOR:
                return (getSubRumors().size() > 0 ? getSubRumors().get(0).toString() : "");
            case LEVEL:
                return (category.equals(RumorCategory.INFLUENCE) && getSubRumors().size() > 0 ?
                        "" + getSubRumors().get(0).getLevel() : "");
            case TYPE:
                if (category.equals(RumorCategory.INFLUENCE)) {
                    return (getSubRumors().size() > 0 ?
                            "Level " + getSubRumors().get(0).getLevel() + " " + this.getName() :
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


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public LocalDate getDate() {
        return date;
    }

    public void setDate(LocalDate date) {
        this.date = date;
    }

    public String getEvent() {
        return event;
    }

    public void setEvent(String event) {
        this.event = event;
    }

    public String getSubEvent() {
        return subEvent;
    }

    public void setSubEvent(String subEvent) {
        this.subEvent = subEvent;
    }

    public int getLevel() {
        return level;
    }

    public void setLevel(int level) {
        this.level = level;
    }

    public List<Rumor> getSubRumors() {
        return subRumors;
    }

    public void setSubRumors(List<Rumor> subRumors) {
        this.subRumors = subRumors;
    }

    public void addSubRumor(Rumor subRumor) {
        subRumors.add(subRumor);
    }

}
