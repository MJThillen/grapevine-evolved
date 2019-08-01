package grapevine;

import grapevine.constants.Constants;

import java.time.LocalDate;
import java.util.List;

public class Event {

    public enum EVENT_TYPE {
        none(0),
        action(1),
        plot(2),
        rumor(3),
        cause(4),
        effect(5);


        private int value;

        EVENT_TYPE(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    private String name; // action, cause, effect - for c&e: item
    private String event; // action, rumor, plot, cause, effect - for c&e: subitem
    private String result; // action
    private EVENT_TYPE eventType; //new+all

    private int level; // action, rumor
    private int unused; // action
    private int total; // action
    private int growth; // action

    private List<Event> causes; // action, rumor, plot
    private List<Event> effects; // action, plot
    private List<String> recipients; //rumor
    private LocalDate date; //plot

    public Event() {

    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEvent() {
        return event;
    }

    public void setEvent(String event) {
        this.event = event;
    }

    public String getResult() {
        return result;
    }

    public void setResult(String result) {
        this.result = result;
    }

    public int getLevel() {
        return level;
    }

    public void setLevel(int level) {
        this.level = level;
    }

    public int getUnused() {
        return unused;
    }

    public void setUnused(int unused) {
        this.unused = unused;
    }

    public int getTotal() {
        return total;
    }

    public void setTotal(int total) {
        this.total = total;
    }

    public int getGrowth() {
        return growth;
    }

    public void setGrowth(int growth) {
        this.growth = growth;
    }

    public List<Event> getCauses() {
        return causes;
    }

    public void setCauses(List<Event> causes) {
        this.causes = causes;
    }

    public List<Event> getEffects() {
        return effects;
    }

    public void setEffects(List<Event> effects) {
        this.effects = effects;
    }

    public EVENT_TYPE getEventType() {
        return eventType;
    }

    public void setEventType(EVENT_TYPE eventType) {
        this.eventType = eventType;
    }

    public List<String> getRecipients() {
        return recipients;
    }

    public void setRecipients(List<String> recipients) {
        this.recipients = recipients;
    }

    public LocalDate getDate() {
        return date;
    }

    public void setDate(LocalDate date) {
        this.date = date;
    }

    public String shortDesc(LocalDate when) {
        if (!when.equals(date)) {
            return date + " " + this.toString();
        } else {
            return this.toString();
        }
    }

    public String toString() {
        String base = "";
        switch(eventType) {
            case none:
                return base;
            case action:
                return name.trim() + " " + event.trim() + " Action";
            case plot:
                return name.trim() + " Plot";
            case rumor:
                if (event.isEmpty()) {
                    return name.trim() + " Rumor";
                } else if (name.endsWith(" Influence")) {
                    return name.substring(0, name.length()-9) + event + " Rumor";
                } else {
                    return name.trim() + " " + event.trim() + " Rumor";
                }
            case cause:
                return "Cause: " + name.trim() + " on " + date;
            case effect:
                return "Effect: " + name.trim() + " on " + date;
        }
        return base;
    }

}
